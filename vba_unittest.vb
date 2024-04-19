Option Compare Database
Option Explicit
Dim TotalTest_Failed As Long, TotalTest_Passed As Long

Sub AssertNumberEquals(value1 As Long, value2 As Long)
    If value1 <> value2 Then
        Debug.Print "Failed: expected " & value2 & ", got " & value1 & "."
        TotalTest_Failed = TotalTest_Failed + 1
    Else
        Debug.Print "Passed"
        TotalTest_Passed = TotalTest_Passed + 1
    End If
End Sub

Sub AssertTextEquals(value1 As String, value2 As String)
    If value1 <> value2 Then
        Debug.Print "Failed: expected " & value2 & ", got " & value1 & "."
        TotalTest_Failed = TotalTest_Failed + 1
    Else
        Debug.Print "Passed"
        TotalTest_Passed = TotalTest_Passed + 1
    End If
End Sub

Sub AssertBoolEquals(value1 As Boolean, value2 As Boolean)
    If value1 <> value2 Then
        Debug.Print "Failed: expected " & value2 & ", got " & value1 & "."
        TotalTest_Failed = TotalTest_Failed + 1
    Else
        Debug.Print "Passed"
        TotalTest_Passed = TotalTest_Passed + 1
    End If
End Sub


Sub TestNumToCN()
    AssertTextEquals NumberToColumnName(1), "A"
    AssertTextEquals NumberToColumnName(26), "Z"
    AssertTextEquals NumberToColumnName(27), "AA"
    AssertTextEquals NumberToColumnName(53), "BA"
    AssertTextEquals NumberToColumnName(702), "ZZ"
    AssertTextEquals NumberToColumnName(676), "ZZ"
End Sub

Sub ResetTest()
    TotalTest_Failed = 0
    TotalTest_Passed = 0
End Sub

Sub PlayResult()
    Debug.Print (TotalTest_Failed + TotalTest_Passed) & " tests conducted, " & TotalTest_Failed & " failed."
End Sub

Private Sub RunTests()
    ResetTest
    
    TestNumToCN
    ' Call other unit test sub procedures here
    PlayResult
    ResetTest
    
End Sub

