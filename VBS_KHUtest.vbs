'**************************************************************************

Const MainCirclePosition_X    = 55.6 '55.6		
Const MainCirclePosition_Y    = 54.95 '54.95		
Const MainCircleRadius_R  		 = 48.4'47.78'47.6
Const ProductAmount			      = 25
Const Z_AXIS 							 = 17
Const ProductHight					 = 80.5'83'82.65 '82.5'85
Const MarkingOrientation		 = 0


'**************************************************************************
Const Debug_msgbox_active    	= 0
Const lmosMORootId						= 100

Const Pi    = 3.1415926535897932384626433832795
'*************************************************************************

Sub LaserMarker_ScriptBegin ()

  Dim mo_Data
  Dim mo_fact_id
	Dim mo_Matrix
  Dim mo_MainCircleRadius
  Dim sp
  Dim sp_id
  Dim sp_matrix
  Dim wrad
  Dim Angle
  Dim dActPos
	Dim Myaxis
  Dim PositionOffsetX(25)
  Dim PositionOffsetY(25)
  Dim PositionEnabled(25)
	Dim MatrixOffsetX (25)
	Dim MatrixOffsetY(25)
	Dim TmpVarX(25)
	Dim TmpVarY(25)
	Dim last_digit
	
	TmpVarX(1)= 1.0
	TmpVarX(2)= 1.0
	TmpVarX(3)= 1.0
  	TmpVarX(4)= 1.0
  	TmpVarX(5)= 1.0
  	TmpVarX(6)= 1.0
  	TmpVarX(7)= 1.0
  	TmpVarX(8)= 1.0
  	TmpVarX(9)= 1.0
  	TmpVarX(10)= 1.0
  	TmpVarX(11)= 1.0
  	TmpVarX(12)= 1.15
  	TmpVarX(13)= 0.5
  	TmpVarX(14)= 1.0
  	TmpVarX(15)= 0.55
  	TmpVarX(16)= 0.5
  	TmpVarX(17)= 0.7
  	TmpVarX(18)= 1.0
  	TmpVarX(19)= 1.0
  TmpVarX(20)= 1.0
  TmpVarX(21)= 1.0
  	TmpVarX(22)= 1.0
	TmpVarX(23)= 1.0
  	TmpVarX(24)= 1.0
  	TmpVarX(25)= 1.0
      
	TmpVarY(1)= 1.0
	TmpVarY(2)= 1.0
	TmpVarY(3)= 1.0
  	TmpVarY(4)= 1.0
  	TmpVarY(5)= 1.0
  	TmpVarY(6)= 1.0
  	TmpVarY(7)= 1.0
  	TmpVarY(8)= 1.0
  	TmpVarY(9)= 1.0
  	TmpVarY(10)= 1.0
  	TmpVarY(11)= 1.0
  	TmpVarY(12)= 1.0
  	TmpVarY(13)= 1.0
  	TmpVarY(14)= 1.0
  	TmpVarY(15)= 1.0
  	TmpVarY(16)= 1.0
  	TmpVarY(17)= 1.0
  	TmpVarY(18)= 1.4
  	TmpVarY(19)= 0.3
  TmpVarY(20)= 1.0
  TmpVarY(21)= 1.0
  	TmpVarY(22)= 1.0
	TmpVarY(23)= 1.0
  	TmpVarY(24)= 1.0
  	TmpVarY(25)= 1.0
      

	PositionEnabled(1) = True'False
	PositionEnabled(2) = True
	PositionEnabled(3) = True
	PositionEnabled(4) = True
	PositionEnabled(5) = True
	PositionEnabled(6) = True
	PositionEnabled(7) = True
	PositionEnabled(8) = False'True
	PositionEnabled(9) = False'True
	PositionEnabled(10) = True
	PositionEnabled(11) = True
	PositionEnabled(12) = True
	PositionEnabled(13) = True
	PositionEnabled(14) = True
	PositionEnabled(15) = True
	PositionEnabled(16) = True
	PositionEnabled(17) = True
	PositionEnabled(18) = True
	PositionEnabled(19) = True
	PositionEnabled(20) = False'True
	PositionEnabled(21) = True
	PositionEnabled(22) = True
	PositionEnabled(23) = True
	PositionEnabled(24) = True
	PositionEnabled(25) = True

	MatrixOffsetX(1) = -0.55
	MatrixOffsetX(2) = -0.1
	MatrixOffsetX(3) = 0.2
	MatrixOffsetX(4) = 0.5
	MatrixOffsetX(5) = 0.7
	MatrixOffsetX(6) = 1.1
	MatrixOffsetX(7) = 0.90
	MatrixOffsetX(8) = 1.05 '1.05
	MatrixOffsetX(9) = 1 '1
	MatrixOffsetX(10) = 0.8 '0.8
	MatrixOffsetX(11) = 0.8'0.8 '0.6
	MatrixOffsetX(12) = 0.35 '0.35
	MatrixOffsetX(13) = -0.15
	MatrixOffsetX(14) = -0.20 '-0.3
	MatrixOffsetX(15) = -0.6
	MatrixOffsetX(16) = -0.8'-0.85
	MatrixOffsetX(17) = -1.2'-1.1
	MatrixOffsetX(18) = -1.6
	MatrixOffsetX(19) = -1.8 '-1.95
	MatrixOffsetX(20) = -2.20'-2.15'-1.65'-2.1 '-2.1
	MatrixOffsetX(21) = -1.77'-1.8'-1.91 '-1.9
	MatrixOffsetX(22) = -1.75'-1.7 '-1.8 '.4
	MatrixOffsetX(23) = -1.55'-1.45 '-1.7 '.4
	MatrixOffsetX(24) = -1.25'-1.15 '-1.4 '.4
	MatrixOffsetX(25) = -0.9'-0.7 '-1.1 '.4

	MatrixOffsetY(1) = -1.25
	MatrixOffsetY(2) = -1.35 
	MatrixOffsetY(3) = -1.4
	MatrixOffsetY(4) = -1.1
	MatrixOffsetY(5) = -0.85
	MatrixOffsetY(6) = -0.55
	MatrixOffsetY(7) = -0.25
	MatrixOffsetY(8) = 0.1 '0
	MatrixOffsetY(9) = 0.4 '0.3
	MatrixOffsetY(10) = 0.6'0.8 '0.6
	MatrixOffsetY(11) = 0.75 '0.9
	MatrixOffsetY(12) = 0.9'1.1 '0.9
	MatrixOffsetY(13) = 1
	MatrixOffsetY(14) = 0.95
	MatrixOffsetY(15) = 0.9
	MatrixOffsetY(16) = 0.85'0.7
	MatrixOffsetY(17) = 0.7'0.6
	MatrixOffsetY(18) = 0.7
	MatrixOffsetY(19) = -0.05 '-0.1
	MatrixOffsetY(20) = -0.5'-0.7'-0.6 '-0.3
	MatrixOffsetY(21) = -0.5'-0.3 '-0.8 
	MatrixOffsetY(22) = -0.7'-1.0 '-0.8
	MatrixOffsetY(23) = -0.9'-1.0'-1.1 '-0.9
	MatrixOffsetY(24) = -1.05 '-1.15'-1.25 '-1.15
	MatrixOffsetY(25) = -1.2 '-1.2


  PositionOffsetX(1) = -0.4
  PositionOffsetY(1) = 0.05
  PositionOffsetX(2) = -0.3
  PositionOffsetY(2) = -0.10
  PositionOffsetX(3) = -0.3
  PositionOffsetY(3) = -0.10
  PositionOffsetX(4) = -0.3 'zmniejszony dystans
  PositionOffsetY(4) = 0'-0.15
  PositionOffsetX(5) = -0.35
  PositionOffsetY(5) = -0.05
  PositionOffsetX(6) = -0.2
  PositionOffsetY(6) = -0.05
  PositionOffsetX(7) = -0.4
  PositionOffsetY(7) = -0.05
  PositionOffsetX(8) = -0.3 '-0.3
  PositionOffsetY(8) = -0.3 '0.01
  PositionOffsetX(9) = -0.2 '-0.2
  PositionOffsetY(9) = -0.5 '-0.01
  PositionOffsetX(10) = -0.3 '-0.3
  PositionOffsetY(10) = -0.1
  PositionOffsetX(11) = -0.15 '-0.15
  PositionOffsetY(11) = -0.2
  PositionOffsetX(12) = -0.3 '-0.3
  PositionOffsetY(12) = -0.2

  PositionOffsetX(13) = -0.3
  PositionOffsetY(13) = -0.2
  PositionOffsetX(14) = -0.2 'zmniejszony dystans
  PositionOffsetY(14) = -0.35 '-0.3
  PositionOffsetX(15) = -0.4
  PositionOffsetY(15) = -0.3
  PositionOffsetX(16) = -0.4
  PositionOffsetY(16) = -0.25
  PositionOffsetX(17) = -0.5
  PositionOffsetY(17) = -0.2
  PositionOffsetX(18) = -0.5'-0.6
  PositionOffsetY(18) = -0.2
  PositionOffsetX(19) = -0.6 '-0.7
  PositionOffsetY(19) = -0.25
  PositionOffsetX(20) = -0.8 '-0.8
  PositionOffsetY(20) = -0.25
  PositionOffsetX(21) = -0.6'-0.7 '-1.1 '-0.7
  PositionOffsetY(21) =  -0.2'-0.3'-0.1 '-0.2
  PositionOffsetX(22) = -0.55'-0.75 '-0.75
  PositionOffsetY(22) = -0.0'-0.15
  PositionOffsetX(23) = -0.55'-0.75 '0.75
  PositionOffsetY(23) = 0.10'-0.0'-0.1
  PositionOffsetX(24) = -0.45'-0.45 '-0.65
  PositionOffsetY(24) = 0.10'0.10 '-0.05
  PositionOffsetX(25) = -0.45'-0.55 '-0.75
  PositionOffsetY(25) = 0.05


  If Debug_msgbox_active=1 Then
		Set mo_MainCircleRadius = lasermarker.drawing.getmosbyname("MainCircleRadius").item(1)
		Set sp = mo_MainCircleRadius.sizepos
		Set MyAxis = LaserMarker.Axis
		sp.x = MainCirclePosition_X
		sp.y = MainCirclePosition_Y
		sp.width = 2*MainCircleRadius_R
		sp.height = sp.width
		mo_MainCircleRadius.sizepos = sp
	Else 
		Set mo_MainCircleRadius = lasermarker.drawing.getmosbyname("MainCircleRadius").item(1)
		Set sp = mo_MainCircleRadius.sizepos
		Set MyAxis = LaserMarker.Axis
		sp.x = 0.1
		sp.y = 0.1
		sp.width = 0.1
		sp.height = sp.width
		mo_MainCircleRadius.sizepos = sp
  End If

	
	Set mo_Data = lasermarker.drawing.getmosbyname("Data").item(1)
  Set mo_Fact_id = lasermarker.drawing.getmosbyname("Kod").item(1)
	Set mo_Matrix = lasermarker.drawing.getmosbyname("Matrix").item(1)

	Set sp = mo_Data.sizepos
  Set sp_id = mo_Fact_id.sizepos
	Set sp_matrix = mo_Matrix.sizepos

	Set MyAxis = LaserMarker.Axis
	
	Hight = MyAxis.GetPos(Z_AXIS, True)
	If Hight<> ProductHight Then
		MyAxis.NewPos Z_AXIS, ProductHight,True
		MyAxis.MoveAxes
	End If

		Dim text1	
		Dim i
		Dim k
		text1 = mo_Data
			For i = 1 To ProductAmount	
				Read_final_serial_from_txt()
				generate_last()
				Angle = ((360/ProductAmount)*i) - (360/ProductAmount )
				wrad = Angle/180*Pi 
	
				sp.x = 	cos (wrad)*(MainCircleRadius_R+PositionOffsetX(i))+MainCirclePosition_X-sin(wrad)*PositionOffsetY(i)
				sp.y = 	sin (wrad)*(MainCircleRadius_R+PositionOffsetX(i))+MainCirclePosition_Y+cos(wrad)*PositionOffsetY(i)
				sp.angle= Angle+MarkingOrientation

				sp_matrix.x = cos (wrad)*(MainCircleRadius_R+PositionOffsetX(i)+0.5)+MainCirclePosition_X + MatrixOffsetX(i) -PositionOffsetX(i)-sin(wrad)*(PositionOffsetY(i)+TmpVarX(i))
				sp_matrix.y = 	sin (wrad)*(MainCircleRadius_R+PositionOffsetX(i)+0.5)+MainCirclePosition_Y + MatrixOffsetY(i) -PositionOffsetY(i)+cos(wrad)*(PositionOffsetY(i)+TmpVarY(i))
				sp_matrix.angle = Angle+MarkingOrientation

				sp_id.x = 	cos (wrad)*(MainCircleRadius_R+PositionOffsetX(i)-2.02)+MainCirclePosition_X-sin(wrad)*PositionOffsetY(i)
				sp_id.y = 	sin (wrad)*(MainCircleRadius_R+PositionOffsetX(i)-2.02)+MainCirclePosition_Y+cos(wrad)*PositionOffsetY(i)
				sp_id.angle= Angle+MarkingOrientation

				lasermarker.scriptutils.message "MarkingObjekt = " & i
				lasermarker.scriptutils.message "Angle ° = " & Angle
				lasermarker.scriptutils.message "Pos x = " & sp.x
				lasermarker.scriptutils.message "Pos y = " & sp.y
				'lasermarker.scriptutils.message "  "
				lasermarker.scriptutils.message " Data "	&Mo_Data.Value 



				Mo_Data.Value = str(serial_doc_duoc + 1) + "H" + last_digit
          Mo_Matrix.Value = Right("000000" & (text1 + k), 6) + mo_Fact_id.Value + "K"

				
				k = k + 1


				mo_Data.sizepos = sp
			  'mo_Fact_id.sizepos = sp_id
				mo_Matrix.sizepos = sp_matrix

				If PositionEnabled(i) = True Then mo_Data.mark
			  'If PositionEnabled(i) = True Then mo_Fact_id.mark
				If PositionEnabled(i) = True Then mo_Matrix.mark	

				lasermarker.synchronize
	
					If Debug_msgbox_active=1 Then
						msgbox "Press F5 in MarkingArea Window" & vbCrLf & " Loop Number: " & i & vbCrLf & "Select 'OK' to continue"
					End If
			Next		
			Save_lastserial()
Function generate_last()
	for ()
		if eru1 < 71 || eru1 > 77
	

	last_digit = 123456
End Function


End Sub
