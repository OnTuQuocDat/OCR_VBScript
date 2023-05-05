Dim Fixture_Rows
Fixture_Rows = 8
Dim Fixture_Column
Fixture_Column = 10
Dim lastLine
lastLine = "BB3ZR0D"
Dim result_index
Dim save_serial_fornext
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
Dim generate
Dim ts
Set ts = fso.OpenTextFile("C:\Users\DAQ\Downloads\Job_Gamma_VBcode\OCR_VBScript\test_generate_serial.txt")
generate = Split(ts.ReadAll, vbCrLf)
ts.Close
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
    WScript.Echo "Hello World!"
    Read_final_serial_from_txt()
    find_index()
    'Start loop for i for j ngay đây
    For i = 0 To Fixture_Rows - 1
        For j = 0 To Fixture_Column - 1
            Generate_serial()
            mo_Data = "S" + value_khac_serial
            WScript.Echo mo_Data
        Next
    Next
    Save_lastserial()

End Sub

Function find_index()
    ' Tim index cua ký tự chuỗi lastLine  đang nằm ở index thứ mấy trong file tổng --> sau đó lấy "result index" ra để cộng thêm 80 con tiếp theo
    Dim text
    Dim ts
    Set ts = fso.OpenTextFile("C:\Users\DAQ\Downloads\Job_Gamma_VBcode\OCR_VBScript\test_generate_serial.txt")
    text = Split(ts.ReadAll, vbCrLf)
    ts.Close
    result_index = -1
    For i = 0 To UBound(text)
        If text(i) = lastLine Then
            result_index = i
            Exit For
        End If
    Next
    If result_index >= 0 Then
        WScript.Echo result_index
    Else
        WScript.Echo "Cannot find the index for this serial in file"
    End If
End Function

Function Read_final_serial_from_txt()
    'Đọc chuỗi string cuối cùng trong file temp.txt
    Dim ts
    Set ts = fso.OpenTextFile("C:\Users\DAQ\Downloads\Job_Gamma_VBcode\OCR_VBScript\temp.txt")
    Dim lines
    lines = Split(ts.ReadAll, vbCrLf)
    ts.Close
    lastLine = lines(UBound(lines) - 1)
    WScript.Echo lastLine
End Function

Function Generate_serial()
    'For i As Integer = result_index To result_index + 81
    '   Console.WriteLine("i value: {0}", i)
    'Next
    ''''''''''''''''''''''''''Note chỗ này vô nè
    If count < 81 Then
        value_khac_serial = generate(result_index + count)
        count = count + 1
    End If
    ''''''''''''''''''''
End Function

Function Save_lastserial()
    'Find index cua value_khac_serial
    Dim index_khac
    index_khac = -1
    For i = 0 To UBound(generate)
        If generate(i) = value_khac_serial Then
            index_khac = i
            Exit For
        End If
    Next
    save_serial_fornext = generate(index_khac)
    'save_serial_fornext = "6W3NV4W"
    Dim file_save_name
    file_save_name = "C:\Users\DAQ\Downloads\Job_Gamma_VBcode\OCR_VBScript\temp.txt"
    Dim objWriter
    Set objWriter = fso.OpenTextFile(file_save_name, 8, True)
    objWriter.WriteLine save_serial_fornext
    objWriter.Close
    WScript.Echo "Write done"
End Function



