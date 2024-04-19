Option Compare Database
Option Explicit
'https://codekabinett.com/rdumps.php?Lang=2&targetDoc=windows-api-declaration-vba-64-bit
Private Declare PtrSafe Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal wRevert As Long) As Long
Private Declare PtrSafe Function EnableMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long

Public Function IsSystemTable(tableName As String) As Boolean
    IsSystemTable = False
    
    If left(tableName, 4) = "MSys" Then
        IsSystemTable = True
    End If
End Function

Public Function DropTableIfExists(table_name As String) As Boolean
    Dim db As DAO.Database
    Dim tbl As DAO.TableDef
    
    ' Open the current database
    Set db = CurrentDb
    
    DropTableIfExists = False
    ' Check if the table exists
    For Each tbl In db.TableDefs
        If tbl.Name = table_name Then
            ' Drop the table
            db.Execute "DROP TABLE " & table_name
            DropTableIfExists = True
            Exit Function
        End If
    Next tbl
    
    ' Clean up
    Set tbl = Nothing
    Set db = Nothing
End Function

Public Function TableExist(tbl_name As String) As Boolean
    Dim db As DAO.Database
    Dim tbl As DAO.TableDef
    
    ' Open the current database
    Set db = CurrentDb
    
    TableExist = False
    ' Check if the table exists
    For Each tbl In db.TableDefs
        If tbl.Name = tbl_name Then
            TableExist = True
        End If
    Next tbl
    
    ' Clean up
    Set tbl = Nothing
    Set db = Nothing
End Function

Public Function SysLoginToEmail(ByVal login As String) As String
    Dim db As Database, rs As Recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset("tbl_cbk_sys_users", dbOpenSnapshot, dbReadOnly)
    If rs.EOF Then
        SysLoginToEmail = ""
        GoTo Exit_Handler
    End If
    rs.FindFirst "LCase(login)='" & LCase(login) & "'"
    If rs.NoMatch Then
        SysLoginToEmail = ""
        GoTo Exit_Handler
    Else
        SysLoginToEmail = rs!email
    End If
Exit_Handler:
    rs.Close
    Set rs = Nothing
    db.Close
    Set db = Nothing
End Function


Public Function RunScheduledTask() As String
    Dim db As Database, rs As Recordset
    Dim current_user_id As String, rs_Str As String, Instruction As String

    current_user_id = FindGUID(TempVars![ID])
    Set db = CurrentDb
    Set rs = db.OpenRecordset("tbl_cbk_scheduledTask", dbOpenSnapshot, dbReadOnly)
    If rs.EOF Then
        RunScheduledTask = ""
        'Debug.Print "EOF"
        GoTo Exit_Handler
    End If
    rs_Str = "scheduleAccount=" & current_user_id & _
        " AND DATEDIFF('n',scheduleTime, #" & _
        Format(Now, "dd-MMM-yyyy hh:mm:ss") & "#) = 0"
    rs.FindFirst rs_Str
    If rs.NoMatch Then
        RunScheduledTask = ""
        'Debug.Print rs_Str & ": No Match"
    Else
        Instruction = CStr(rs!scheduleTask)
        RunScheduledTask = Instruction
        'Debug.Print rs_Str & ": " & Instruction
    End If
Exit_Handler:
    rs.Close
    Set rs = Nothing
    db.Close
    Set db = Nothing
    
End Function

Static Function Log10(X)
    Log10 = Log(X) / Log(10#)
End Function

Public Sub Initialization()
'https://stackoverflow.com/questions/34901810/microsoft-access-loop-through-all-forms-and-controls-on-each-form
    Dim FrmName As String
    Dim Frm As Form
    Dim i As Integer, stay_open As Boolean
    
    stay_open = False
    Application.Echo False
    SysCmd acSysCmdInitMeter, "Initialization...", CurrentProject.AllForms.Count
    For i = 0 To CurrentProject.AllForms.Count - 1
        FrmName = CurrentProject.AllForms(i).FullName
        If InStr(1, FrmName, "sub") > 0 And Not FrmName = "subfrm_cbk_inventory_management" Then
            If Not CurrentProject.AllForms(i).IsLoaded Then
                'DoCmd.OpenForm FrmName, acDesign, , , , acHidden
                DoCmd.OpenForm FrmName, acNormal, , , , acHidden
            Else
                stay_open = True
            End If
            Set Frm = Forms(FrmName)
            InitForm Frm
            If Frm.RecordSource <> "" Then
                Debug.Print Format(FrmName & ".RecordSource", "!" & String(74, "@")), "[" & Forms(FrmName).RecordSource & "]"
                'Forms(FrmName).RecordSource = ""
                Frm.RecordSource = ""
                Debug.Print "After clearing: ", "[" & Forms(FrmName).RecordSource & "]"
            Else
                Debug.Print Format(FrmName & ".RecordSource", "!" & String(74, "@")), "Ok"
            End If
            If Not stay_open Then
                DoCmd.Close acForm, FrmName, acSaveYes
                stay_open = False
            End If
        End If
        SysCmd acSysCmdUpdateMeter, i
    Next i
    'Debug.Print "Done"
    SysCmd acSysCmdRemoveMeter
    Application.Echo True
End Sub

Public Sub InitForm(ByVal ofrm As Form)
    Dim Ctrl As Control
    Dim CtrlType As String, FrmName As String
    
    FrmName = ofrm.Name
    For Each Ctrl In ofrm.Controls
        'Debug.Print Ctrl.ControlType, VarType(Ctrl.ControlType)
        If Ctrl.ControlType = acComboBox Or Ctrl.ControlType = acTextBox Then
            If Ctrl.ControlSource <> "" Then
                Debug.Print String(4, " ") & Format(FrmName & "." & Ctrl.Name & ".ControlSource", "!" & String(70, "@")), "[" & Ctrl.ControlSource & "]"
                Ctrl.ControlSource = ""
                Debug.Print String(4, " ") & "After clearing: ", "[" & Ctrl.ControlSource & "]"
            Else
                Debug.Print String(4, " ") & Format(FrmName & "." & Ctrl.Name & ".ControlSource", "!" & String(70, "@")), "Ok"
            End If
        End If
    Next
End Sub

Public Function SQLByDate(dt_str As String, fieldname As String) As String
    Dim next_day As Date
    Dim fmt_str As String
    If Not IsDate(dt_str) Then
        SQLByDate = ""
        GoTo Exit_Handler
    End If
    'MyPrint "Original:  " & dt_str
    fmt_str = Format(dt_str, "YYYY-MM-DD HH:NN:SS")
    'MyPrint "Processed: " & fmt_str
    next_day = DateValue(Format(dt_str, "YYYY-MM-DD HH:NN:SS")) + 1
    'MyPrint next_day
    fmt_str = fieldname & ">=#" & Format(dt_str, "YYYY-MM-DD HH:NN:SS") & _
        "# AND " & fieldname & "<#" & Format(next_day, "YYYY-MM-DD HH:NN:SS") & "#"
    
    'MyPrint fmt_str
    SQLByDate = fieldname & ">=#" & Format(dt_str, "YYYY-MM-DD HH:NN:SS") & _
        "# AND " & fieldname & "<#" & Format(next_day, "YYYY-MM-DD HH:NN:SS") & "#"
Exit_Handler:
    Exit Function
End Function

Public Function WhereInArray(ObjArray As Variant, ObjToFind As Variant) As Integer
'DEVELOPER: Ryan Wells (wellsr.com)
'DESCRIPTION: Function to check where a value is in an array
    Dim i As Long
    For i = LBound(ObjArray) To UBound(ObjArray)
        If StringLimitedCompare(ObjArray(i), ObjToFind) Then
            WhereInArray = i + 1
            Exit Function
        End If
    Next i
    'if you get here, ObjToFind was not in the array. Set to null
    WhereInArray = -1
End Function

Public Function PrintDictionary(ByVal dic As Dictionary, Optional prettyprint As Boolean = False) As String
' Advanced datatype in VBA, to add "Microsoft Scripting Runtime" in References
    Dim k As Variant
    Dim retn_str As String
    
    If TypeName(dic) <> "Dictionary" Then
        PrintDictionary = ""
        Exit Function
    End If
    
    For Each k In dic.Keys
        If prettyprint Then
            If retn_str = "" Then
                retn_str = Format(DateTime.Now, "yyyy-MM-dd hh:mm:ss") & vbTab & CStr(k) & " : " & dic(k)
            Else
                retn_str = retn_str & vbCrLf & _
                Format(DateTime.Now, "yyyy-MM-dd hh:mm:ss") & vbTab & CStr(k) & " : " & dic(k)
            End If
        Else
            If retn_str = "" Then
                retn_str = CStr(k) & " : " & dic(k)
            Else
                retn_str = retn_str & ", " & CStr(k) & " : " & dic(k)
            End If
        End If
    Next k
    PrintDictionary = retn_str
End Function

Public Function StringLimitedCompare(ByVal Subject As Variant, ByVal reference As Variant, Optional ByVal case_sensitive As Boolean = False) As Boolean
    Dim ref As String, subj As String
    Dim i As Integer
    StringLimitedCompare = True
    If Len(subj) < Len(ref) Then
        StringLimitedCompare = False
        Exit Function
    End If
    ref = CStr(reference)
    subj = CStr(Subject)
    For i = 1 To Len(ref)
        If case_sensitive Then
            'Debug.Print Mid(subj, i, 1) & " : " & Mid(ref, i, 1) & "; Case-sensitive"
            If Asc(Mid(subj, i, 1)) <> Asc(Mid(ref, i, 1)) Then
                StringLimitedCompare = False
                Exit Function
            End If
            'Debug.Print "ok"
        Else
            'Debug.Print LCase(Mid(subj, i, 1)) & " : " & LCase(Mid(ref, i, 1))
            If LCase(Mid(subj, i, 1)) <> LCase(Mid(ref, i, 1)) Then
                StringLimitedCompare = False
                Exit Function
            End If
            'Debug.Print "ok"
        End If
    Next i
End Function

Public Function isGUID(ByVal strGUID As String) As Boolean
' Credits to https://stackoverflow.com/questions/133493/check-for-a-valid-guid
    Dim regEx As RegExp
    Set regEx = New RegExp
    If IsNull(strGUID) Or strGUID = "" Then
        isGUID = False
        Exit Function
    End If
    regEx.Pattern = "^({|\()?[A-Fa-f0-9]{8}-([A-Fa-f0-9]{4}-){3}[A-Fa-f0-9]{12}(}|\))?$"
    isGUID = regEx.Test(strGUID)
    Set regEx = Nothing
End Function

Public Function FileExists(ByVal path_ As String) As Boolean
    ' Credits to https://stackoverflow.com/questions/44434199/access-vba-check-if-file-exists
    FileExists = (Len(Dir(path_)) > 0)
End Function

Public Function MachineName() As String
    MachineName = Environ$("computername")
End Function

'' Validate email address
Public Function ValidateEmailAddress(ByVal strEmailAddress As String) As Boolean
    'Credits to https://www.geeksengine.com/article/validate-email-vba.html
    'On Error GoTo Error_Handler
    
    Dim objRegExp As New RegExp
    Dim blnIsValidEmail As Boolean
    
    ValidateEmailAddress = False
    objRegExp.IgnoreCase = True
    objRegExp.Global = True
    objRegExp.Pattern = "^([a-zA-Z0-9_\-\.]+)@[a-z0-9-]+(\.[a-z0-9-]+)*(\.[a-z]{2,3})$"
    
    blnIsValidEmail = objRegExp.Test(strEmailAddress)
    ValidateEmailAddress = blnIsValidEmail
      
    Exit Function
    
Error_Handler:
    ValidateEmailAddress = False
    'MsgBox "Module: - ValidateEmailAddress function" & vbCrLf & vbCrLf _
        & "Error#:  " & Err.Number & vbCrLf & vbCrLf & Err.Description
End Function

Public Sub ExportData()
    On Error GoTo Error_Handler
    ' To use Excel VBA, must set the reference library
    'https://stackoverflow.com/questions/5729195/how-to-refer-to-excel-objects-in-access-vba
    Dim cfg_rs As Recordset, rs As Recordset
    Dim filename As String, sql_string As String, sht_name As String, qry_name As String
    Dim xlApp As Excel.Application
    Dim targetWorkbook As Excel.Workbook
    Dim targetWkSht As Excel.Worksheet
    Dim i As Integer, j As Integer
    Dim dict As Dictionary, db As Object

    DoCmd.Hourglass True
    Application.Echo False
    
    Set db = CurrentDb()
    filename = TempVars![Hotel name] & "_Export Data_" & Format(DateTime.Now, "yyyyMMdd_hhmm") & ".xlsx"
    MyPrint filename
    Set xlApp = New Excel.Application
    xlApp.EnableEvents = False
    'xlApp.DisplayAlerts = False
    xlApp.Visible = True
    Set targetWorkbook = xlApp.Workbooks.Add
    
    Set dict = New Dictionary
    dict.Add "tbl_cbk_admin_log", "qry_cbk_exp_admin_log"
    dict.Add "tbl_cbk_inventory_management", "qry_cbk_exp_inventory_management"
    dict.Add "tbl_cbk_guest_transaction", "qry_cbk_exp_guest_transaction"
    dict.Add "tbl_cbk_room_inventory", "qry_cbk_exp_room_inventory"
    dict.Add "tbl_cbk_hotel_guests", "qry_cbk_exp_hotel_guests"
    'dict.Add "tbl_cbk_sys_config", "qry_cbk_exp_sys_config"
    'dict.Add "tbl_cbk_sys_users", "qry_cbk_exp_sys_users"

    For j = 0 To dict.Count - 1
        sht_name = dict.Keys(j)
        qry_name = dict.items(j)
        If j + 1 > targetWorkbook.Sheets.Count Then
            Set targetWkSht = targetWorkbook.Sheets.Add
        Else
            Set targetWkSht = targetWorkbook.Sheets(j + 1)
        End If
        targetWkSht.Name = sht_name

        Set rs = db.OpenRecordset(qry_name)
        For i = 0 To rs.Fields.Count - 1
            If HasProperty(rs.Fields(i), "Caption") Then
                targetWkSht.Cells(1, i + 1).Value = rs.Fields(i).Properties("Caption").Value
            Else
                targetWkSht.Cells(1, i + 1).Value = rs.Fields(i).Name
            End If
        Next
        targetWkSht.Range("A2").CopyFromRecordset rs
        targetWkSht.ListObjects.Add(, targetWkSht.UsedRange, , xlYes).Name = sht_name
        'MyPrint sht_name & ": " & targetWkSht.UsedRange.Address
    Next

    targetWorkbook.SaveAs filename
    
    xlApp.Quit
    Set xlApp = Nothing
    cfg_rs.Close
    rs.Close
    db.Close
    Set cfg_rs = Nothing
    Set rs = Nothing
    Set db = Nothing
    Set dict = Nothing
    MsgBox "Data exported to" & vbCrLf & filename, vbInformation, "Data export"

Exit_Handler:
    On Error Resume Next
    DoCmd.Hourglass False
    Application.Echo True
    Exit Sub
 
Error_Handler:
    xlApp.EnableEvents = False
    xlApp.Quit
        
    MsgBox "MS Access has generated the following error" & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Source: btn_import_data_Click" & vbCrLf & _
           "Error Description: " & Err.Description, _
           vbCritical, "An Error has Occurred!"
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    Resume Exit_Handler

End Sub

Public Function RandBetween(ByVal lower_range As Long, ByVal upper_range As Long) As Long
    Dim c_range As Long
    c_range = upper_range - lower_range
    RandBetween = (CLng(upper_range - lower_range + 1) * Rnd + lower_range) Mod c_range + lower_range
End Function

Public Function FindGUID(ByVal strGUID As String) As String
'Returns the valid GUID if found in a given string can tolerate upper/lowercase.
'Returns null if no valid GUID
'Same tolerance with Microsoft own GUIDFromString

    Dim regEx As RegExp
    Dim mat As Object
    Set regEx = New RegExp
    If IsNull(strGUID) Then
        FindGUID = ""
        Exit Function
    End If
    'regEx.Pattern = "^({|\()?[A-Fa-f0-9]{8}-([A-Fa-f0-9]{4}-){3}[A-Fa-f0-9]{12}(}|\))?$"
    regEx.Pattern = "({|\()+[A-Fa-f0-9]{8}-([A-Fa-f0-9]{4}-){3}[A-Fa-f0-9]{12}(}|\))"
    Set mat = regEx.Execute(strGUID)
    If mat.Count = 0 Then
        FindGUID = ""
    Else
        FindGUID = UCase(mat(0))
    End If
    Set regEx = Nothing
End Function

Function RandString(n As Long) As String
'Credits to https://stackoverflow.com/questions/34643520/how-to-generate-a-string-of-random-characters-from-a-given-list-of-characters
'Don't re-invent the wheel
    'Assumes that Randomize has been invoked by caller
    Dim i As Long, j As Long, m As Long, s As String, pool As String
    pool = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
    m = Len(pool)
    Randomize
    For i = 1 To n
        j = 1 + Int(m * Rnd())
        s = s & Mid(pool, j, 1)
    Next i
    RandString = s
End Function

Public Function GenerateSalt(username As String, password As String, Optional uppercase As Boolean = True) As String
    Dim salt As String
    
    If uppercase Then
        salt = SHA256(UCase(username) & Chr(255) & password)
    Else
        salt = SHA256(username & Chr(255) & password)
    End If
    GenerateSalt = salt
End Function

Public Function PasswordRule() As String
    Dim rule_string As String
    Dim lcl As Integer, ucl As Integer, num As Integer, sym As Integer, pwl As Integer
    
    rule_string = ""
    
    If IsNull(TempVars![Password letter]) Then
        lcl = 0
    Else
        lcl = TempVars![Password letter]
    End If
    
    If IsNull(TempVars![Password cap]) Then
        ucl = 0
    Else
        ucl = TempVars![Password cap]
    End If
    
    If IsNull(TempVars![Password number]) Then
        num = 0
    Else
        num = TempVars![Password number]
    End If
    
    If IsNull(TempVars![Password symbol]) Then
        sym = 0
    Else
        sym = TempVars![Password symbol]
    End If
    
    If IsNull(TempVars![Password length]) Then
        pwl = 0
    Else
        pwl = TempVars![Password length]
    End If
    
    If lcl > 0 Then
        If lcl > 1 Then
            rule_string = rule_string & vbCrLf & vbTab & lcl & " lower case letters or more needed"
        Else
            rule_string = rule_string & vbCrLf & vbTab & lcl & " lower case letter or more needed"
        End If
    Else
        rule_string = rule_string & vbCrLf & vbTab & "Lower case letter not required"
    End If

    If ucl > 0 Then
        If ucl > 1 Then
            rule_string = rule_string & vbCrLf & vbTab & ucl & " upper case letters or more needed"
        Else
            rule_string = rule_string & vbCrLf & vbTab & ucl & " upper case letter or more needed"
        End If
    Else
        rule_string = rule_string & vbCrLf & vbTab & "Upper case letter not required"
    End If

    If num > 0 Then
        If num > 1 Then
            rule_string = rule_string & vbCrLf & vbTab & num & " numbers or more needed"
        Else
            rule_string = rule_string & vbCrLf & vbTab & num & " number or more needed"
        End If
    Else
        rule_string = rule_string & vbCrLf & vbTab & "Number not required"
    End If

    If sym > 0 Then
        If sym > 1 Then
            rule_string = rule_string & vbCrLf & vbTab & sym & " symbols or more needed"
        Else
            rule_string = rule_string & vbCrLf & vbTab & sym & " symbol or more needed"
        End If
    Else
        rule_string = rule_string & vbCrLf & vbTab & "Symbol not required"
    End If

    rule_string = rule_string & vbCrLf & vbTab & "Password length : " & pwl

    PasswordRule = rule_string
End Function

'If you want to disable the SHIFT key, type ap_DisableShift in the Immediate window,
'and then press ENTER. If you want to enable the shift key,
'type ap_EnableShift in the Immediate window, and then press ENTER.

Public Sub ap_DisableShift()
'https://docs.microsoft.com/en-us/office/troubleshoot/access/disable-database-startup-options
'This function disable the shift at startup. This action causes
'the Autoexec macro and Startup properties to always be executed.

    On Error GoTo errDisableShift
    
    Dim db As DAO.Database
    Dim prop As DAO.Property
    Const conPropNotFound = 3270
    
    Set db = CurrentDb()
    
    'This next line disables the shift key on startup.
    db.Properties("AllowByPassKey") = False
    
    'The function is successful.


exitDisableShift:
    db.Close
    Set db = Nothing
    Exit Sub

errDisableShift:
'The first part of this error routine creates the "AllowByPassKey
'property if it does not exist.
    If Err = conPropNotFound Then
        Set prop = db.CreateProperty("AllowByPassKey", _
        dbBoolean, False)
        db.Properties.Append prop
        Resume Next
    Else
        MsgBox "Function 'ap_DisableShift' did not complete successfully."
    End If
    GoTo exitDisableShift

End Sub

Public Sub ap_EnableShift()
'This function enables the SHIFT key at startup. This action causes
'the Autoexec macro and the Startup properties to be bypassed
'if the user holds down the SHIFT key when the user opens the database.

    On Error GoTo errEnableShift
    
    Dim db As DAO.Database
    Dim prop As DAO.Property
    Const conPropNotFound = 3270
    
    Set db = CurrentDb()
    
    'This next line of code disables the SHIFT key on startup.
    db.Properties("AllowByPassKey") = True
    
    'function successful
    Debug.Print "Shift Key enabled"
exitEnableShift:
    db.Close
    Set db = Nothing
    Exit Sub

errEnableShift:
'The first part of this error routine creates the "AllowByPassKey
'property if it does not exist.
    If Err = conPropNotFound Then
        Set prop = db.CreateProperty("AllowByPassKey", _
        dbBoolean, True)
        db.Properties.Append prop
        Resume Next
    Else
        MsgBox "Function 'ap_DisableShift' did not complete successfully."
    End If
    GoTo exitEnableShift
End Sub

Public Function HasTable(tbl_name As String) As Boolean
' https://stackoverflow.com/questions/2985513/check-if-access-table-exists/2992743#2992743
' Find a table with matching name in the CurrentDb
' Return Nothing if unable to find the table

    Dim db As DAO.Database
    Dim td As DAO.TableDef
    Set db = CurrentDb
    On Error Resume Next
    Set td = db.TableDefs(tbl_name)
    HasTable = (Err.Number = 0)
    Err.Clear

End Function

Public Function HasField(tbl_name As String, fld_name As String) As Boolean
' Find a field with matching name in a given table
' Return Nothing if unable to find the field
    Dim db As Object, tdf As TableDef, fld As Field
    
    Set db = CurrentDb
    HasField = False
    
    For Each tdf In db.TableDefs
        If tdf.Name = tbl_name Then
            For Each fld In tdf.Fields
                If fld.Name = fld_name Then
                    HasField = True
                    Exit For
                End If
            Next fld
        End If
    Next tdf
    db.Close
    Set db = Nothing
End Function

Public Sub DisplayTempVars()
' Traverse all the TempVars and print the name, value and type of data
    Dim i As Integer, vtype As String
    
    For i = 0 To TempVars.Count - 1
        Select Case VarType(TempVars(i).Value)
            Case vbEmpty
                vtype = "Empty"
            Case vbNull
                vtype = "Null"
            Case vbInteger
                vtype = "Integer"
            Case vbLong
                vtype = "Long"
            Case vbDouble
                vtype = "Double"
            Case vbDate
                vtype = "Date"
            Case vbString
                vtype = "String"
            Case vbBoolean
                vtype = "Boolean"
            Case Else
                vtype = CStr(VarType(TempVars(i).Value))
        End Select
        Debug.Print "TempVars![" & _
            Format(TempVars(i).Name & "]", "!" & String(20, "@")) & " is (" & _
            Format(vtype & ")", "!@@@@@@@@@@") & "[" & TempVars(i).Value & "]"
    Next
End Sub

Public Function IsAccde() As Boolean
'https://stackoverflow.com/questions/41724465/how-to-detect-if-your-microsoft-access-application-is-compiled-i-e-running-as
    On Error Resume Next

    IsAccde = False
    'IsAccde = (InStr(CurrentDb.Properties("Name"), "accde") = 0)
    IsAccde = (CurrentDb.Properties("MDE") = "T")
End Function

Public Sub LoadSysConfig()
    On Error GoTo Error_Handler
    'Load all name-value pair in tbl_cbk_config_sys into TempVars
    Dim db As Database, rs As Recordset
    Dim i As Integer
    
    DoCmd.Hourglass True
    Application.Echo False
    Set db = CurrentDb()
    Set rs = db.OpenRecordset("qry_cbk_sys_config", dbOpenSnapshot, dbReadOnly)
    rs.MoveFirst
    Do Until rs.EOF
        If Not IsNull(TempVars.Item(CStr(rs!sys_config_name))) Then
            TempVars.Remove (rs!sys_config_name)
        End If
        Select Case rs!sys_config_type
            Case "double"
                If IsNull(rs!sys_config_value) Then
                    TempVars.Add rs!sys_config_name, 0
                Else
                    TempVars.Add rs!sys_config_name, CDbl(rs!sys_config_value)
                End If
            
            Case "integer"
                If IsNull(rs!sys_config_value) Then
                    TempVars.Add rs!sys_config_name, 0
                Else
                    TempVars.Add rs!sys_config_name, CInt(rs!sys_config_value)
                End If
            
            Case "date"
                If IsNull(rs!sys_config_value) Then
                    TempVars.Add rs!sys_config_name, CDate("1/1/1900")
                Else
                    TempVars.Add rs!sys_config_name, CDate(rs!sys_config_value)
                End If
            
            Case "boolean"
                If IsNull(rs!sys_config_value) Then
                    TempVars.Add rs!sys_config_name, False
                Else
                    If StrConv(Trim(CStr(rs!sys_config_value)), vbProperCase) = "True" Then
                        TempVars.Add rs!sys_config_name, True
                    Else
                        TempVars.Add rs!sys_config_name, False
                    End If
                End If
            
            Case Else
                If IsNull(rs!sys_config_value) Then
                    TempVars.Add rs!sys_config_name, ""
                Else
                    TempVars.Add rs!sys_config_name, CStr(rs!sys_config_value)
                End If
        
        End Select
        rs.MoveNext
    Loop

Exit_Handler:
    On Error Resume Next
    rs.Close
    db.Close
    Set rs = Nothing
    Set db = Nothing
    DoCmd.Hourglass False
    Application.Echo True
    Exit Sub
 
Error_Handler:
    MsgBox "MS Access has generated the following error" & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Source: btn_import_data_Click" & vbCrLf & _
           "Error Description: " & Err.Description, _
           vbCritical, "An Error has Occurred!"
    'Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    Resume Exit_Handler
    
End Sub

Public Sub AccessStatus(message As String)
    Dim varReturn As Variant
    varReturn = SysCmd(acSysCmdSetStatus, message)
End Sub

Public Sub AccessStatusClr()
    SysCmd (acSysCmdClearStatus)
End Sub

Public Sub MyPrint(ParamArray str() As Variant)
    Dim final_str As String
    Dim i As Integer, j As Integer
    final_str = ""
    For i = LBound(str) To UBound(str)
        Select Case VarType(str(i))
        Case 8192 To 8200:
            For j = LBound(str(i)) To UBound(str(i))
                final_str = final_str & vbTab & CStr(str(i)(j))
            Next j
        Case Else:
            final_str = final_str & vbTab & CStr(str(i))
        End Select
    Next i
    
    Debug.Print Format(Format(DateTime.Now, "yyyy-MM-dd hh:mm:ss"), String(20, "@")) & _
        Format("  [" & Application.CurrentProject.Name & "]", "!" & String(40, "@")) & final_str
End Sub

Public Function IsUpper(C As String) As Boolean
    'https://bettersolutions.com/vba/strings-characters/ascii-characters.htm
    IsUpper = Asc(C) >= 65 And Asc(C) <= 90
End Function

Public Function IsLower(C As String) As Boolean
    IsLower = Asc(C) >= 97 And Asc(C) <= 122
End Function

Public Function IsAlpha(C As String) As Boolean
    IsAlpha = IsUpper(C) Or IsLower(C)
End Function

Public Function IsLoginAllowableC(C As String) As Boolean
    'Asc("@") = 64: Asc(".") = 46: Asc("_") = 95
    IsLoginAllowableC = IsAlpha(C) Or IsNumeric(C) Or Asc(C) = 64 Or Asc(C) = 46 Or Asc(C) = 95
End Function

Public Function IsLoginAllowableS(C As String) As Boolean
    Dim i As Integer
    
    For i = 1 To Len(C)
        If IsLoginAllowableC(Mid(C, i, 1)) Then
            IsLoginAllowableS = True
        Else
            IsLoginAllowableS = False
            Exit For
        End If
    Next
End Function

Public Function InBetw(num As Integer, lr As Integer, ur As Integer) As Boolean
    If ur < lr Then
        Err.Raise 17, "mod_cbk_common.InBetw", "Variable ur (lpper range) should be higher than lr (lower range)"
    End If
    InBetw = num <= ur And num >= lr
End Function

Public Function IsSpecialChar(C As String) As Boolean
    Dim a As Integer
    a = Asc(C)
    IsSpecialChar = InBetw(a, 33, 47) Or InBetw(a, 58, 64) Or InBetw(a, 91, 96) Or InBetw(a, 123, 126)
End Function

Public Function HasNumber(C As String) As Boolean
    Dim i As Integer
    HasNumber = False
    For i = 1 To Len(C)
        If IsNumeric(Mid(C, i, 1)) Then
            HasNumber = True
            Exit For
        End If
    Next
End Function

Public Function HasLower(C As String) As Boolean
    Dim i As Integer
    HasLower = False
    For i = 1 To Len(C)
        If IsLower(Mid(C, i, 1)) Then
            HasLower = True
            Exit For
        End If
    Next
End Function

Public Function HasUpper(C As String) As Boolean
    Dim i As Integer
    HasUpper = False
    For i = 1 To Len(C)
        If IsUpper(Mid(C, i, 1)) Then
            HasUpper = True
            Exit For
        End If
    Next
End Function

Public Function HasSpecialChar(C As String) As Boolean
    Dim i As Integer
    HasSpecialChar = False
    For i = 1 To Len(C)
        If IsSpecialChar(Mid(C, i, 1)) Then
            HasSpecialChar = True
            Exit For
        End If
    Next
End Function

Public Function LenSymbol(C As String) As Integer
    Dim i As Integer
    LenSymbol = 0
    For i = 1 To Len(C)
        If IsSpecialChar(Mid(C, i, 1)) Then
            LenSymbol = LenSymbol + 1
        End If
    Next
End Function

Public Function LenUpper(C As String) As Integer
    Dim i As Integer
    LenUpper = 0
    For i = 1 To Len(C)
        If IsUpper(Mid(C, i, 1)) Then
            LenUpper = LenUpper + 1
        End If
    Next
End Function

Public Function LenLower(C As String) As Integer
    Dim i As Integer
    LenLower = 0
    For i = 1 To Len(C)
        If IsLower(Mid(C, i, 1)) Then
            LenLower = LenLower + 1
        End If
    Next
End Function

Public Function LenNumber(C As String) As Integer
    Dim i As Integer
    LenNumber = 0
    For i = 1 To Len(C)
        If IsNumeric(Mid(C, i, 1)) Then
            LenNumber = LenNumber + 1
        End If
    Next
End Function

Public Function HasProperty(fld As Field, pty As String) As Boolean
    Dim i As Integer
    HasProperty = False
    For i = 0 To fld.Properties.Count - 1
        If fld.Properties(i).Name = pty Then
            HasProperty = True
            Exit For
        End If
    Next
End Function

Public Function ExecuteSQLwStatus(sql_string As String) As Integer
    On Error GoTo Error_Handler
    ExecuteSQLwStatus = 0
    CurrentDb.Execute sql_string, dbFailOnError
    
Exit_Handler:
    On Error Resume Next
    Exit Function
 
Error_Handler:
    ExecuteSQLwStatus = Err.Number
    Resume Exit_Handler

End Function

Public Sub sendOutlookEmail()
    'https://stackoverflow.com/questions/17973549/ms-access-vba-sending-an-email-through-outlook
    'https://www.encodedna.com/excel/how-to-parse-outlook-emails-and-show-in-excel-worksheet-using-vba.htm
    Dim oApp As Outlook.Application
    Dim oMail As MailItem
    Set oApp = CreateObject("Outlook.application")

    Set oMail = oApp.CreateItem(olMailItem)
    oMail.body = "Body of the email"
    oMail.Subject = "Test Subject"
    oMail.To = "boon_kwee_chan_from.csog@moh.gov.sg"
    'oMail.To = "rehmen_izwan_jailani@moh.gov.sg"
    'oMail.SentOnBehalfOfName = "GMS@moh.gov.sg"    'invalid Sender email
    oMail.Send
    
    Set oMail = Nothing
    Set oApp = Nothing

End Sub

Public Sub GMSSendEmail(addressee_email_addr As String, message As String, subj As String, Optional on_behalf As String = "")
    'https://stackoverflow.com/questions/17973549/ms-access-vba-sending-an-email-through-outlook
    'https://www.encodedna.com/excel/how-to-parse-outlook-emails-and-show-in-excel-worksheet-using-vba.htm
    Dim oApp As Outlook.Application
    Dim oMail As MailItem
    Set oApp = CreateObject("Outlook.application")

    Set oMail = oApp.CreateItem(olMailItem)
    oMail.body = message
    oMail.Subject = subj
    oMail.To = addressee_email_addr
    If Len(on_behalf) > 1 Then
        oMail.SentOnBehalfOfName = on_behalf
    End If
    oMail.Send
    Set oMail = Nothing
    Set oApp = Nothing

End Sub


Public Function ConvToBase64String(vIn As Variant) As Variant
    'https://en.wikibooks.org/wiki/Visual_Basic_for_Applications/String_Hashing_in_VBA
    'Check that Net Framework 3.5 (includes .Net 2 and .Net 3 is installed in windows
    'and not just Net Advanced Services
   
   Dim oD As Object
      
    Set oD = CreateObject("MSXML2.DOMDocument")
    With oD
        .LoadXML "<root />"
        .DocumentElement.DataType = "bin.base64"
        .DocumentElement.nodeTypedValue = vIn
    End With
    ConvToBase64String = Replace(oD.DocumentElement.Text, vbLf, "")
    
    Set oD = Nothing

End Function

Public Function ConvToHexString(vIn As Variant) As Variant
    'https://en.wikibooks.org/wiki/Visual_Basic_for_Applications/String_Hashing_in_VBA
    
    Dim oD As Object
      
    Set oD = CreateObject("MSXML2.DOMDocument")
      
      With oD
        .LoadXML "<root />"
        .DocumentElement.DataType = "bin.Hex"
        .DocumentElement.nodeTypedValue = vIn
      End With
    ConvToHexString = Replace(oD.DocumentElement.Text, vbLf, "")
    
    Set oD = Nothing

End Function

Public Function SHA256(sIn As String, Optional bB64 As Boolean = 0) As String
    'https://en.wikibooks.org/wiki/Visual_Basic_for_Applications/String_Hashing_in_VBA
    'Set a reference to mscorlib 4.0 64-bit
    'SHA256("password")
    'same output as digest::digest(algo = c("sha256"), object = "password", serialize = FALSE) in R
    
    'Test with empty string input:
    '64 Hex:   e3b0c44298f...etc
    '44 Base-64:   47DEQpj8HBSa+/...etc
    
    Dim oT As Object, oSHA256 As Object
    Dim TextToHash() As Byte, bytes() As Byte
    
    Set oT = CreateObject("System.Text.UTF8Encoding")
    Set oSHA256 = CreateObject("System.Security.Cryptography.SHA256Managed")
    
    TextToHash = oT.GetBytes_4(sIn)
    bytes = oSHA256.ComputeHash_2((TextToHash))
    
    If bB64 = True Then
       SHA256 = ConvToBase64String(bytes)
    Else
       SHA256 = ConvToHexString(bytes)
    End If
    
    Set oT = Nothing
    Set oSHA256 = Nothing
    
End Function

'https://www.fmsinc.com/microsoftaccess/startup/preventclose.asp
Public Sub AccessCloseButtonEnabled(pfEnabled As Boolean)
  ' Comments: Control the Access close button.
  '           Disabling it forces the user to exit within the application
  ' Params  : pfEnabled       TRUE enables the close button, FALSE disabled it
  ' Owner   : Copyright (c) FMS, Inc.
  ' Source  : Total Visual SourceBook
  ' Usage   : Permission granted to subscribers of the FMS Newsletter

  On Error Resume Next

  Const clngMF_ByCommand As Long = &H0&
  Const clngMF_Grayed As Long = &H1&
  Const clngSC_Close As Long = &HF060&

  Dim lngWindow As Long
  Dim lngMenu As Long
  Dim lngFlags As Long

  lngWindow = Application.hWndAccessApp
  lngMenu = GetSystemMenu(lngWindow, 0)
  If pfEnabled Then
    lngFlags = clngMF_ByCommand And Not clngMF_Grayed
  Else
    lngFlags = clngMF_ByCommand Or clngMF_Grayed
  End If
  Call EnableMenuItem(lngMenu, clngSC_Close, lngFlags)
End Sub

Public Function RegExpTest(strString As String, strRegExp As String, Optional bolIgnoreCase As Boolean = False) As Boolean
    Dim re As Object
    Set re = CreateObject("vbscript.RegExp")
    re.Pattern = strRegExp
    re.IgnoreCase = bolIgnoreCase
    RegExpTest = re.Test(strString)
    Set re = Nothing
End Function

Public Function RegExpReplace(strString As String, strRegExp As String, _
    Optional repl_str As String = "", Optional bolIgnoreCase As Boolean = False) As String
    
    Dim re As Object
    Set re = CreateObject("vbscript.RegExp")
    
    re.Global = True
    re.Pattern = strRegExp
    re.IgnoreCase = bolIgnoreCase
    RegExpReplace = re.Replace(strString, repl_str)
    Set re = Nothing
End Function

Public Function NricReGexTest(ByVal nric As String) As Boolean
    Dim objRegExp As New RegExp
    
    NricReGexTest = False
    objRegExp.IgnoreCase = True
    objRegExp.Global = True
    objRegExp.Pattern = "^[mstgf]{1}[0-9]{7}[abcdefghijklmnpqrtuwxz]{1}$"
    
    NricReGexTest = objRegExp.Test(LCase(nric))
End Function

Public Function NRIC_Checksum_Test(ByVal nric As String) As Boolean
    Dim objRegExp As New RegExp
    Dim total As Integer, l_nric As String, first_letter As String, last_letter As String
    Dim xtra As Integer
    
    l_nric = LCase(nric)
    first_letter = Mid(l_nric, 1, 1)
    last_letter = Mid(l_nric, 9, 1)
    
    'objRegExp.Global = True
    objRegExp.Pattern = "^[mstgf]{1}\d{7}[abcdefghijklmnpqrtuwxz]{1}$"
    
    If objRegExp.Test(l_nric) Then
        If InStr(1, "tg", first_letter) > 0 Then
            xtra = 4
        ElseIf InStr(1, "m", first_letter) > 0 Then
            xtra = 3
        Else
            xtra = 0
        End If
        total = (CInt(Mid(nric, 2, 1)) * 2) + _
                (CInt(Mid(nric, 3, 1)) * 7) + _
                (CInt(Mid(nric, 4, 1)) * 6) + _
                (CInt(Mid(nric, 5, 1)) * 5) + _
                (CInt(Mid(nric, 6, 1)) * 4) + _
                (CInt(Mid(nric, 7, 1)) * 3) + _
                (CInt(Mid(nric, 8, 1)) * 2) + xtra
        
        Select Case total Mod 11
            Case 0
                If InStr(1, "gfm", first_letter) > 0 Then
                    NRIC_Checksum_Test = last_letter = "x"
                Else
                    NRIC_Checksum_Test = last_letter = "j"
                End If
            Case 1
                If InStr(1, "gfm", first_letter) > 0 Then
                    NRIC_Checksum_Test = last_letter = "w"
                Else
                    NRIC_Checksum_Test = last_letter = "z"
                End If
            Case 2
                If InStr(1, "gfm", first_letter) > 0 Then
                    NRIC_Checksum_Test = last_letter = "u"
                Else
                    NRIC_Checksum_Test = last_letter = "i"
                End If
            Case 3
                If InStr(1, "gfm", first_letter) > 0 Then
                    NRIC_Checksum_Test = last_letter = "t"
                Else
                    NRIC_Checksum_Test = last_letter = "h"
                End If
            Case 4
                If InStr(1, "gfm", first_letter) > 0 Then
                    NRIC_Checksum_Test = last_letter = "r"
                Else
                    NRIC_Checksum_Test = last_letter = "g"
                End If
            Case 5
                If InStr(1, "gfm", first_letter) > 0 Then
                    NRIC_Checksum_Test = last_letter = "q"
                Else
                    NRIC_Checksum_Test = last_letter = "f"
                End If
            Case 6
                If InStr(1, "gfm", first_letter) > 0 Then
                    NRIC_Checksum_Test = last_letter = "p"
                Else
                    NRIC_Checksum_Test = last_letter = "e"
                End If
            Case 7
                If InStr(1, "gfm", first_letter) > 0 Then
                    NRIC_Checksum_Test = last_letter = "n"
                Else
                    NRIC_Checksum_Test = last_letter = "d"
                End If
            Case 8
                If InStr(1, "gfm", first_letter) > 0 Then
                    NRIC_Checksum_Test = last_letter = "m"
                Else
                    NRIC_Checksum_Test = last_letter = "c"
                End If
            Case 9
                If InStr(1, "gfm", first_letter) > 0 Then
                    NRIC_Checksum_Test = last_letter = "l"
                Else
                    NRIC_Checksum_Test = last_letter = "b"
                End If
            Case 10
                If InStr(1, "gfm", first_letter) > 0 Then
                    NRIC_Checksum_Test = last_letter = "k"
                Else
                    NRIC_Checksum_Test = last_letter = "a"
                End If
        End Select
    Else
        NRIC_Checksum_Test = False
    End If
End Function

Sub Test_runtest()
    AssertBoolEquals StringLimitedCompare("Test", "Test ", False), False
    AssertBoolEquals StringLimitedCompare("Test ", "Test", False), True
End Sub

Sub RunTestsHere()
    ResetTest
    Test_runtest
    PlayResult
End Sub
