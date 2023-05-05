'a = msgbox("Have a good day, fellow GFG reader!", 0, "Making a VBScript program")


Const Fixture_Column = 10
Const Fixture_Rows = 8
Dim lastLine 
Dim result_index
Dim save_serial_fornext
Dim value_khac_serial 
Dim count

count = 1

Dim i
Dim j
Dim k
Dim text1
Dim mo_Data


Main()

Sub Main()
    find_index()
End Sub

' Function find_index()
'     Dim objFSO,objInputFile,tmpStr,substrtoFind
'     Set objFSO = CreateObject("Scripting.FileSystemObject")
'     substrtoFind = "GVLVB8N"
'     Set objInputFile = objFSO.OpenTextFile("test_generate_serial.txt",2,"True")
'     'tmpStr = objInputFile.ReadLine
'     'result_index = getIndex(tmpStr)
'     result_index = getIndex(substrtoFind)
'     msgbox "Index" & result_index

' End Function

Function find_index()
    'Dim fso, file, text, index,substrtoFind
    Dim substrtoFind
    substrtoFind = "2201ZT6"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.OpenTextFile("C:\Users\DAQ\Downloads\Job_Gamma_VBcode\OCR_VBScript\test_generate_serial.txt",1)
    ' text = file.ReadAll
    ' file.Close

    ' index = InStr(text,substrtoFind)

    ' If index > 0 Then
    '     WScript.Echo "Index of search string: " & index
    ' Else
    '     WScript.Echo "Search string not found"
    ' End If

    Do Until file.AtEndOfStream
        strLine = file.ReadLine
        intIndex = InStr(strLine,substrtoFind)
        If index > 0 Then
            WScript.Echo "Index of 'substrtoFind' in line '" & strLine & "': " & intIndex
        End If
    Loop


    file.Close
End Function

