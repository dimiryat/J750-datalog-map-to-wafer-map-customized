Attribute VB_Name = "Module1"
Option Explicit


'"readFileContents" is to read from text file and store it to a string array for process
Sub readFileContents(ByVal fullFilename As String, ByRef Return_str() As String)

    Dim objFSO As Object
    Dim objTF As Object
    Dim strIn As String

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objTF = objFSO.OpenTextFile(fullFilename, 1)
    strIn = objTF.readall
    objTF.Close
    Return_str = Split(strIn, vbCrLf)
        
End Sub

'"Macro_entry" is this file's entry point
Sub Macro_entry()

    Dim strPattern1, strPattern2 As String
    Dim temp_str As Variant
    Dim Raw_str() As String
    Dim regEx As New RegExp
    Dim dummy_id As Integer: dummy_id = 1
    Dim final_collection As New Collection
    Dim unit_info_collection As New Collection

    readFileContents Worksheets("Source").Range("B1").Value, Raw_str
        
    'Regular expression pattern for site 0
    strPattern1 = " [X,Y] Coordinate byte[0-9] [0-9]+ Site 0  DECIMAL: [0-9]+"
    strPattern2 = "    0       [\s,0-9][\s,0-9][0-9]         [0-9]"
    
    'Sort out site0's information
    For Each temp_str In Raw_str
        With regEx
            .Pattern = strPattern1
            .Global = True
        End With
        If regEx.test(temp_str) Then
            final_collection.Add CStr(dummy_id) & ": " & temp_str
        End If
        
        regEx.Pattern = strPattern2
        If regEx.test(temp_str) Then
            final_collection.Add CStr(dummy_id) & ": " & temp_str
            dummy_id = dummy_id + 1
        End If
    Next temp_str
    
    'Regular expression pattern for site 1
    strPattern1 = " [X,Y] Coordinate byte[0-9] [0-9]+ Site 1  DECIMAL: [0-9]+"
    strPattern2 = "    1       [\s,0-9][\s,0-9][0-9]         [0-9]"
    
    'Sort out site1's information
    For Each temp_str In Raw_str
        With regEx
            .Pattern = strPattern1
            .Global = True
        End With
        If regEx.test(temp_str) Then
            final_collection.Add CStr(dummy_id) & ": " & temp_str
        End If
        
        regEx.Pattern = strPattern2
        If regEx.test(temp_str) Then
            final_collection.Add CStr(dummy_id) & ": " & temp_str
            dummy_id = dummy_id + 1
        End If
    Next temp_str
    
    'Regular expression pattern for input data
    strPattern1 = "([0-9]+):  ([X,Y]) Coordinate byte([0-9]) [0-9]+ Site [0-9]  DECIMAL: ([0-9]+)"
    strPattern2 = "[0-9]+:     ([0-9])       ([\s,0-9][\s,0-9][0-9])         [0-9]"
    
    For Each temp_str In final_collection
        Dim temp_unit_info As New Unit_info
        With regEx
            .Pattern = strPattern1
            .Global = True
        End With
        If regEx.test(temp_str) Then
            temp_unit_info.id = CInt(regEx.Replace(temp_str, "$1"))
            If regEx.Replace(temp_str, "$2") = "X" And regEx.Replace(temp_str, "$3") = "1" Then
                temp_unit_info.x_byte1 = CInt(regEx.Replace(temp_str, "$4"))
            ElseIf regEx.Replace(temp_str, "$2") = "X" And regEx.Replace(temp_str, "$3") = "2" Then
                temp_unit_info.x_byte2 = CInt(regEx.Replace(temp_str, "$4"))
            ElseIf regEx.Replace(temp_str, "$2") = "Y" And regEx.Replace(temp_str, "$3") = "1" Then
                temp_unit_info.y_byte1 = CInt(regEx.Replace(temp_str, "$4"))
            ElseIf regEx.Replace(temp_str, "$2") = "Y" And regEx.Replace(temp_str, "$3") = "2" Then
                temp_unit_info.y_byte2 = CInt(regEx.Replace(temp_str, "$4"))
            End If
        End If
        regEx.Pattern = strPattern2
        If regEx.test(temp_str) Then
            temp_unit_info.site = CInt(regEx.Replace(temp_str, "$1"))
            temp_unit_info.bin = CInt(regEx.Replace(temp_str, "$2"))
            temp_unit_info.cal
            unit_info_collection.Add temp_unit_info
            Set temp_unit_info = New Unit_info
        End If
    Next temp_str
    
    Worksheets("Result").Select
    Range("A1").Select
    For Each temp_unit_info In unit_info_collection
        ActiveCell.Value = temp_unit_info.id
        ActiveCell.Offset(0, 1).Select
        ActiveCell.Value = temp_unit_info.site
        ActiveCell.Offset(0, 1).Select
        ActiveCell.Value = temp_unit_info.x_loc
        ActiveCell.Offset(0, 1).Select
        ActiveCell.Value = temp_unit_info.y_loc
        ActiveCell.Offset(1, -3).Select
    Next temp_unit_info
    
    Dim x_loc, y_loc As Long
    
    Worksheets("Wafer map").Select
    For Each temp_unit_info In unit_info_collection
        If (temp_unit_info.x_loc > 0 And temp_unit_info.y_loc > 0) And (temp_unit_info.x_loc <= 52 And temp_unit_info.y_loc <= 286) Then
            Cells(temp_unit_info.y_loc, temp_unit_info.x_loc).Select
            ActiveCell.Value = temp_unit_info.bin
            If temp_unit_info.bin <> 1 Then
                ActiveCell.Interior.Color = RGB(255, 0, 0)
            Else
                ActiveCell.Interior.Color = RGB(0, 0, 255)
            End If
        End If
    Next temp_unit_info
End Sub

'Here are some example code for testing or playing around...
Sub Example_code()
Attribute Example_code.VB_ProcData.VB_Invoke_Func = " \n14"

    Range("D3").Select
    Selection.Value = 3
    MsgBox Selection.Value
    
End Sub

Sub test()

    Cells(10, 10).Select

End Sub
