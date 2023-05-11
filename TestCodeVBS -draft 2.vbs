Const MainPosition_X    = 15.47
Const MainPosition_Y    = 9.81

Const MainCircleRadius_R  		 = 50
Const Z_AXIS 							 = 17
Const ProductHight					 = 82
Const MarkingOrientation		 = 180

Const Fixture_Rows		= 8
Const Fixture_Columns	= 10
Const Spacing_X				= 11.93
Const Spacing_Y		 		= 	9.98


'**************************************************************************
Const Debug_msgbox_active    	= 0
Const lmosMORootId						= 100

Const Pi    = 3.1415926535897932384626433832795

'Dim Fixture_Rows
'Fixture_Rows = 8
'Dim Fixture_Column
'Fixture_Column = 10
Dim lastLine
lastLine = "BB3ZR0D"
Dim result_index
Dim save_serial_fornext
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
Dim generate
Dim ts
Set ts = fso.OpenTextFile("C:\Users\debgc-vlm\Desktop\OCR_VBScript\test_generate_serial.txt")
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



Sub LaserMarker_ScriptBegin()
    'Dim mo_Data
    Dim mo_MainCircleRadius
    Dim sp
    Dim wrad
    Dim Angle
    Dim dActPos
    Dim Myaxis    

    If Debug_msgbox_active=1 Then
        'Set mo_MainCircleRadius = lasermarker.drawing.getmosbyname("MainCircleRadius").item(1)
        'Set sp = mo_MainCircleRadius.sizepos
        'Set MyAxis = LaserMarker.Axis
        'sp.x = MainCirclePosition_X
        'sp.y = MainCirclePosition_Y
        'sp.width = 2*MainCircleRadius_R
        'sp.height = sp.width
        'mo_MainCircleRadius.sizepos = sp
    Else 
        Set mo_MainCircleRadius = lasermarker.drawing.getmosbyname("Main").item(1)
        Set sp = mo_MainCircleRadius.sizepos
        Set MyAxis = LaserMarker.Axis
        sp.x = 0.1
        sp.y = 0.1
        sp.width = 0.1
        sp.height = sp.width
        mo_MainCircleRadius.sizepos = sp
    End If

	Set mo_Data = lasermarker.drawing.getmosbyname("Data").item(1)
	
	Set sp = mo_Data.sizepos
	
	Set MyAxis = LaserMarker.Axis

	Hight = MyAxis.GetPos(Z_AXIS, True)
	If Hight<> ProductHight Then
		MyAxis.NewPos Z_AXIS, ProductHight,True
		MyAxis.MoveAxes
	End If

	'Dim text1
	'Dim i
	'Dim j
	'Dim k
	Dim test

    'text1 = mo_Data
    'decode(text1)

    k = 0
        'Đọc nội dung bắt đầu trong file temp txt
        Read_final_serial_from_txt()
        find_index()
		For i = 0 To Fixture_Rows - 1
			For j = 0 To Fixture_Columns	- 1
				sp.angle = MarkingOrientation
				sp.x = MainPosition_X + i*Spacing_X
				sp.y = MainPosition_Y + j*Spacing_Y
				
				'Read data from txt file to print Value 
          Generate_serial()
                

				Mo_Data.Value = "S" + value_khac_serial
				k = k + 1

				lasermarker.scriptutils.message " "			 'htn
				lasermarker.scriptutils.message "Ban dang khac san pham o vi tri hang " & i + 1 & " cot " & j + 1'htn
				lasermarker.scriptutils.message "Toa do x = " & sp.x	'htn
				lasermarker.scriptutils.message "Toa do y = " & sp.y	'htn
				lasermarker.scriptutils.message " "	'htn

		
				mo_Data.sizepos = sp
				mo_Data.mark


				Do
				Loop While lasermarker.waitoniobit("ShutterOpen",0,10) = True

				lasermarker.synchronize
				
				If Debug_msgbox_active=1 Then
					msgbox "Press F5 in MarkingArea Window" & vbCrLf & " Loop Number: " & i & vbCrLf & "Select 'OK' to continue"
				End If
			Next		
		Next
		'Save serial for next start
		Save_lastserial()
lasermarker.scriptutils.message "Ban dang khac tong cong " & i * j & " con hang. Qua trinh khac ma da hoan tat"

'Start exe file
Set objShell = CreateObject("WScript.Shell")
intReturnCode = objShell.Run("C:\Sonion\new_main.exe", 1, True) 

if intReturnCode = 0 Then
    lasermarker.scriptutils.message "Completed execute"
Else
    lasermarker.scriptutils.message "Not completed execute"

End Sub


Function find_index()
    ' Tim index cua ký tự chuỗi lastLine  đang nằm ở index thứ mấy trong file tổng --> sau đó lấy "result index" ra để cộng thêm 80 con tiếp theo
    Dim text
    Dim ts
    Set ts = fso.OpenTextFile("C:\Users\debgc-vlm\Desktop\OCR_VBScript\test_generate_serial.txt")
    text = Split(ts.ReadAll, vbCrLf)
    ts.Close
    result_index = -1
    For i = 0 To UBound(text)
        If text(i) = lastLine Then
            result_index = i
            Exit For
        End If
    Next
    'If result_index >= 0 Then
     '   WScript.Echo result_index
    'Else
     '   WScript.Echo "Cannot find the index for this serial in file"
    'End If
End Function

Function Read_final_serial_from_txt()
    'Đọc chuỗi string cuối cùng trong file temp.txt
    Dim ts
    Set ts = fso.OpenTextFile("C:\Users\debgc-vlm\Desktop\OCR_VBScript\temp.txt")
    Dim lines
    lines = Split(ts.ReadAll, vbCrLf)
    ts.Close
    lastLine = lines(UBound(lines) - 1)
    'WScript.Echo lastLine
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
    file_save_name = "C:\Users\debgc-vlm\Desktop\OCR_VBScript\temp.txt"
    Dim objWriter
    Set objWriter = fso.OpenTextFile(file_save_name, 8, True)
    objWriter.WriteLine save_serial_fornext
    objWriter.Close
    'WScript.Echo "Write done"
End Function


