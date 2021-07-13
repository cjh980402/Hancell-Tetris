
Sub Selectarea
'그리거나 지우는 영역 선택
	Set center=cells(cells(3,42), cells(3,43))
	'블록의 기준점
	Select Case cells(3,44)
	Case 1
		If cells(3,45)=1 Then
			range(center.offset(0,-1), center.offset(0,2)).select
		elseif cells(3,45)=2 Then
			range(center.offset(-1,0), center.offset(2,0)).select
		End If
	Case 2
		If cells(3,45)=1 Then
			union(range(center.offset(1,-1), center.offset(1,1)), center.offset(0,-1)).select
		elseif cells(3,45)=2 Then
			union(range(center.offset(-1,1), center.offset(1,1)), center.offset(1,0)).select
		elseif cells(3,45)=3 Then
			union(range(center.offset(-1,-1), center.offset(-1,1)), center.offset(0,1)).select
		elseif cells(3,45)=4 Then
			union(range(center.offset(-1,-1), center.offset(1,-1)), center.offset(-1,0)).select		
		End If
	Case 3
		If cells(3,45)=1 Then
			union(range(center.offset(1,-1), center.offset(1,1)), center.offset(0,1)).select
		elseif cells(3,45)=2 Then
			union(range(center.offset(-1,1), center.offset(1,1)), center.offset(-1,0)).select
		elseif cells(3,45)=3 Then
			union(range(center.offset(-1,-1), center.offset(-1,1)), center.offset(0,-1)).select
		elseif cells(3,45)=4 Then
			union(range(center.offset(-1,-1), center.offset(1,-1)), center.offset(1,0)).select
		End If
	Case 4
		If cells(3,45)=1 Then
			union(center.offset(0,-1), center, center.offset(1,0), center.offset(1,1)).select
		elseif cells(3,45)=2 Then
			union(center, center.offset(1,0), center.offset(-1,1), center.offset(0,1)).select
		End If
	Case 5
		If cells(3,45)=1 Then
			union(center.offset(0,1), center, center.offset(1,0), center.offset(1,-1)).select
		elseif cells(3,45)=2 Then
			union(center, center.offset(-1,0), center.offset(1,1), center.offset(0,1)).select
		End If
	Case 6
		If cells(3,45)=1 Then
			union(range(center.offset(1,-1), center.offset(1,1)), center).select
		elseif cells(3,45)=2 Then
			union(range(center.offset(-1,1), center.offset(1,1)), center).select
		elseif cells(3,45)=3 Then
			union(range(center.offset(-1,-1), center.offset(-1,1)), center).select
		elseif cells(3,45)=4 Then
			union(range(center.offset(-1,-1), center.offset(1,-1)), center).select
		End If
	Case 7
		range(center.offset(0,-1), center.offset(1,0)).select
	End Select
End Sub

Sub drawblock
'블록 그리기
	Dim color(6)
	color(0)=RGB(0, 255, 255)
	color(1)=RGB(101, 101, 255)
	color(2)=RGB(255, 165, 0)
	color(3)=RGB(255, 0, 0)
	color(4)=RGB(0, 255, 0)
	color(5)=RGB(170, 0, 255)
	color(6)=RGB(229, 229, 0)
	
	Call Selectarea
	
	on error resume Next
	Selection.interior.colorindex=xlnone
	Selection.interior.color=color(cells(3,44)-1)
	cells(48,48).select
End Sub

Sub drawother(i, btype)
'i=0이면 다음 블록, i=1이면 보관 블록 그리기
	If btype>0 Then
		rtmp=cells(3,42)
		ctmp=cells(3,43)
		ttmp=cells(3,44)
		turntmp=cells(3,45)
		
		cells(3,44)=btype
		cells(3,45)=1
		If cells(3,44)=1 Then
			cells(3,42)=7+i*8
			cells(3,43)=25
		elseif cells(3,44)=7 Then
			cells(3,42)=6+i*8
			cells(3,43)=26
		else
			cells(3,42)=6+i*8
			cells(3,43)=25
		End If
		
		Call drawblock
		
		cells(3,42)=rtmp
		cells(3,43)=ctmp
		cells(3,44)=ttmp
		cells(3,45)=turntmp
	End If
End Sub

Sub Eraseblock
'블록 지우는 함수
	Call Selectarea
	
	on error resume Next
	Selection.interior.colorindex=xlnone
	cells(48,48).select
End Sub

Function checkleft
'왼쪽 이동 가능한지 검사
	Set center=cells(cells(3,42), cells(3,43))
	Select Case cells(3,44)
	Case 1
		If cells(3,45)=1 Then
			center.offset(0,-2).select
		elseif cells(3,45)=2 Then
			range(center.offset(-1,-1), center.offset(2,-1)).select
		End If
	Case 2
		If cells(3,45)=1 Then
			range(center.offset(0,-2), center.offset(1,-2)).select
		elseif cells(3,45)=2 Then
			union(center.offset(-1,0), center, center.offset(1,-1)).select
		elseif cells(3,45)=3 Then
			union(center.offset(-1,-2), center).select
		elseif cells(3,45)=4 Then
			range(center.offset(-1,-2), center.offset(1,-2)).select		
		End If
	Case 3
		If cells(3,45)=1 Then
			union(center, center.offset(1,-2)).select
		elseif cells(3,45)=2 Then
			union(center, center.offset(1,0), center.offset(-1,-1)).select
		elseif cells(3,45)=3 Then
			range(center.offset(-1,-2), center.offset(0,-2)).select
		elseif cells(3,45)=4 Then
			range(center.offset(-1,-2), center.offset(1,-2)).select
		End If
	Case 4
		If cells(3,45)=1 Then
			union(center.offset(0,-2), center.offset(1,-1)).select
		elseif cells(3,45)=2 Then
			union(center.offset(0,-1), center.offset(1,-1), center.offset(-1,0)).select
		End If
	Case 5
		If cells(3,45)=1 Then
			union(center.offset(0,-1), center.offset(1,-2)).select
		elseif cells(3,45)=2 Then
			union(center.offset(-1,-1), center.offset(0,-1), center.offset(1,0)).select
		End If
	Case 6
		If cells(3,45)=1 Then
			union(center.offset(0,-1), center.offset(1,-2)).select
		elseif cells(3,45)=2 Then
			union(center.offset(-1,0), center.offset(0,-1), center.offset(1,0)).select
		elseif cells(3,45)=3 Then
			union(center.offset(-1,-2), center.offset(0,-1)).select
		elseif cells(3,45)=4 Then
			range(center.offset(-1,-2), center.offset(1,-2)).select
		End If
	Case 7
		range(center.offset(0,-2), center.offset(1,-2)).select
	End Select
	
	If Selection.interior.colorindex=xlnone Then
		checkleft=true
	else
		checkleft=false
	End If
	cells(48,48).select
End Function

Function checkright
'오른쪽 이동 가능한지 검사
	Set center=cells(cells(3,42), cells(3,43))
	Select Case cells(3,44)
	Case 1
		If cells(3,45)=1 Then
			center.offset(0,3).select
		elseif cells(3,45)=2 Then
			range(center.offset(-1,1), center.offset(2,1)).select
		End If
	Case 2
		If cells(3,45)=1 Then
			union(center, center.offset(1,2)).select
		elseif cells(3,45)=2 Then
			range(center.offset(-1,2), center.offset(1,2)).select
		elseif cells(3,45)=3 Then
			range(center.offset(-1,2), center.offset(0,2)).select
		elseif cells(3,45)=4 Then
			union(center, center.offset(1,0), center.offset(-1,1)).select		
		End If
	Case 3
		If cells(3,45)=1 Then
			range(center.offset(0,2), center.offset(1,2)).select
		elseif cells(3,45)=2 Then
			range(center.offset(-1,2), center.offset(1,2)).select
		elseif cells(3,45)=3 Then
			union(center, center.offset(-1,2)).select
		elseif cells(3,45)=4 Then
			union(center, center.offset(-1,0), center.offset(1,1)).select
		End If
	Case 4
		If cells(3,45)=1 Then
			union(center.offset(0,1), center.offset(1,2)).select
		elseif cells(3,45)=2 Then
			union(center.offset(-1,2), center.offset(0,2), center.offset(1,1)).select
		End If
	Case 5
		If cells(3,45)=1 Then
			union(center.offset(0,2), center.offset(1,1)).select
		elseif cells(3,45)=2 Then
			union(center.offset(0,2), center.offset(1,2), center.offset(-1,1)).select
		End If
	Case 6
		If cells(3,45)=1 Then
			union(center.offset(0,1), center.offset(1,2)).select
		elseif cells(3,45)=2 Then
			range(center.offset(-1,2), center.offset(1,2)).select
		elseif cells(3,45)=3 Then
			union(center.offset(-1,2), center.offset(0,1)).select
		elseif cells(3,45)=4 Then
			union(center.offset(-1,0), center.offset(0,1), center.offset(1,0)).select
		End If
	Case 7
		range(center.offset(0,1), center.offset(1,1)).select
	End Select
	
	If Selection.interior.colorindex=xlnone Then
		checkright=true
	else
		checkright=false
	End If
	cells(48,48).select
End Function

Function checkdown
'아래쪽 이동 가능한지 검사
	Set center=cells(cells(3,42), cells(3,43))
	Select Case cells(3,44)
	Case 1
		If cells(3,45)=1 Then
			range(center.offset(1,-1), center.offset(1,2)).select
		elseif cells(3,45)=2 Then
			center.offset(3,0).select
		End If
	Case 2
		If cells(3,45)=1 Then
			range(center.offset(2,-1), center.offset(2,1)).select
		elseif cells(3,45)=2 Then
			range(center.offset(2,0), center.offset(2,1)).select
		elseif cells(3,45)=3 Then
			union(center.offset(0,-1), center, center.offset(1,1)).select
		elseif cells(3,45)=4 Then
			union(center, center.offset(2,-1)).select		
		End If
	Case 3
		If cells(3,45)=1 Then
			range(center.offset(2,-1), center.offset(2,1)).select
		elseif cells(3,45)=2 Then
			union(center, center.offset(2,1)).select
		elseif cells(3,45)=3 Then
			union(center.offset(0,1), center, center.offset(1,-1)).select
		elseif cells(3,45)=4 Then
			range(center.offset(2,-1), center.offset(2,0)).select
		End If
	Case 4
		If cells(3,45)=1 Then
			union(center.offset(2,0), center.offset(2,1), center.offset(1,-1)).select
		elseif cells(3,45)=2 Then
			union(center.offset(2,0), center.offset(1,1)).select
		End If
	Case 5
		If cells(3,45)=1 Then
			union(center.offset(2,-1), center.offset(2,0), center.offset(1,1)).select
		elseif cells(3,45)=2 Then
			union(center.offset(1,0), center.offset(2,1)).select
		End If
	Case 6
		If cells(3,45)=1 Then
		range(center.offset(2,-1), center.offset(2,1)).select
		elseif cells(3,45)=2 Then
			union(center.offset(1,0), center.offset(2,1)).select
		elseif cells(3,45)=3 Then
			union(center.offset(0,-1), center.offset(1,0), center.offset(0,1)).select
		elseif cells(3,45)=4 Then
			union(center.offset(2,-1), center.offset(1,0)).select
		End If
	Case 7
		range(center.offset(2,-1), center.offset(2,0)).select
	End Select
	
	If Selection.interior.colorindex=xlnone or Selection.interior.pattern=xlpatterndown Then
		checkdown=true
		'그림자랑 겹칠수도 있으므로 그림자 패턴에도 true반환
	else
		checkdown=false
	End If
	cells(48,48).select
End Function

Function checkrclockturn
'반시계 방향 회전 가능한지 검사
	If cells(3,42)=3 Then
		Call godown
	End If
	
	Set center=cells(cells(3,42), cells(3,43))
	Select Case cells(3,44)
	Case 1
		If cells(3,45)=1 Then
		union(center.offset(1,0), center.offset(2,0), center.offset(-1,0)).select
		elseif cells(3,45)=2 Then
			If cells(3,43)=3 Then
				Call goright
			else
				For i=1 To (cells(3,43)-16)
					Call goleft
				Next
			End If
			union(center.offset(0,1), center.offset(0,2), center.offset(0,-1)).select
		End If
	Case 2
		If cells(3,45)=1 Then
			range(center.offset(-1,1), center.offset(0,1)).select
		elseif cells(3,45)=2 Then
			If cells(3,43)=3 Then
				Call goright
			End If
			range(center.offset(-1,-1), center.offset(-1,0)).select
		elseif cells(3,45)=3 Then
			range(center.offset(0,-1), center.offset(1,-1)).select
		elseif cells(3,45)=4 Then
			If cells(3,43)=18 Then
				Call goleft
			End If
			range(center.offset(1,0), center.offset(1,1)).select		
		End If
	Case 3
		If cells(3,45)=1 Then
			range(center.offset(-1,0), center.offset(-1,1)).select
		elseif cells(3,45)=2 Then
			If cells(3,43)=3 Then
				Call goright
			End If
			range(center.offset(-1,-1), center.offset(0,-1)).select
		elseif cells(3,45)=3 Then
			range(center.offset(1,-1), center.offset(1,0)).select
		elseif cells(3,45)=4 Then
			If cells(3,43)=18 Then
				Call goleft
			End If
			range(center.offset(0,1), center.offset(1,1)).select
		End If
	Case 4
		If cells(3,45)=1 Then
			range(center.offset(-1,1), center.offset(0,1)).select
		elseif cells(3,45)=2 Then
			If cells(3,43)=3 Then
				Call goright
			End If
			union(center.offset(0,-1), center.offset(1,1)).select
		End If
	Case 5
		If cells(3,45)=1 Then
			union(center.offset(-1,0), center.offset(1,1)).select
		elseif cells(3,45)=2 Then
			If cells(3,43)=3 Then
				Call goright
			End If
			range(center.offset(1,-1), center.offset(1,0)).select
		End If
	Case 6
		If cells(3,45)=1 Then
			range(center.offset(-1,1), center.offset(0,1)).select
		elseif cells(3,45)=2 Then
			If cells(3,43)=3 Then
				Call goright
			End If
			range(center.offset(-1,-1), center.offset(-1,0)).select
		elseif cells(3,45)=3 Then
			range(center.offset(0,-1), center.offset(1,-1)).select
		elseif cells(3,45)=4 Then
			If cells(3,43)=18 Then
				Call goleft
			End If
			range(center.offset(1,0), center.offset(1,1)).select
		End If
	End Select
	
	If Selection.interior.colorindex=xlnone or Selection.interior.pattern=xlpatterndown Then
		checkrclockturn=true
		'그림자랑 겹칠수도 있으므로 그림자 패턴에도 true반환
	else
		checkrclockturn=false
	End If
	cells(48,48).select
End Function

Function checkclockturn
'시계 방향 회전 가능한지 검사
	If cells(3,42)=3 Then
		Call godown
	End If
	
	Set center=cells(cells(3,42), cells(3,43))
	Select Case cells(3,44)
	Case 1,4,5
		checkclockturn=checkrclockturn
		Exit Function
	Case 2
		If cells(3,45)=1 Then
			range(center.offset(-1,-1), center.offset(-1,0)).select
		elseif cells(3,45)=2 Then
			If cells(3,43)=3 Then
				Call goright
			End If
			range(center.offset(0,-1), center.offset(1,-1)).select
		elseif cells(3,45)=3 Then
			range(center.offset(1,0), center.offset(1,1)).select
		elseif cells(3,45)=4 Then
			If cells(3,43)=18 Then
				Call goleft
			End If
			range(center.offset(-1,1), center.offset(0,1)).select		
		End If
	Case 3
		If cells(3,45)=1 Then
			range(center.offset(-1,-1), center.offset(0,-1)).select
		elseif cells(3,45)=2 Then
			If cells(3,43)=3 Then
				Call goright
			End If
			range(center.offset(1,-1), center.offset(1,0)).select
		elseif cells(3,45)=3 Then
			range(center.offset(0,1), center.offset(1,1)).select
		elseif cells(3,45)=4 Then
			If cells(3,43)=18 Then
				Call goleft
			End If
			range(center.offset(-1,1), center.offset(-1,0)).select
		End If
	Case 6
		If cells(3,45)=1 Then
			range(center.offset(-1,-1), center.offset(0,-1)).select
		elseif cells(3,45)=2 Then
			If cells(3,43)=3 Then
				Call goright
			End If
			range(center.offset(1,-1), center.offset(1,0)).select
		elseif cells(3,45)=3 Then
			range(center.offset(0,1), center.offset(1,1)).select
		elseif cells(3,45)=4 Then
			If cells(3,43)=18 Then
				Call goleft
			End If
			range(center.offset(-1,0), center.offset(-1,1)).select
		End If
	End Select
	
	If Selection.interior.colorindex=xlnone or Selection.interior.pattern=xlpatterndown Then
		checkclockturn=true
		'그림자랑 겹칠수도 있으므로 그림자 패턴에도 true반환
	else
		checkclockturn=false
	End If
	cells(48,48).select
End Function

Sub hold
	application.screenupdating=false
	
	If cells(5,45)=0 Then '블록 보관하기
		Call Eraseshadow
		Call Eraseblock
		
		For i=0 To 1
			range(cells(4+8*i,23), cells(9+8*i,28)).interior.colorindex=xlnone
		Next
		
		'현재 블록 보관하기
		cells(5,45)=cells(3,44)
		Call drawother(1,cells(5,45))
		
		'다음 블록 등장시키기
		cells(3,42)=2
		cells(3,43)=10
		cells(3,44)=cells(5,44)
		cells(3,45)=1
		cells(5,44)=Int(Rnd*7)+1
		Call drawother(0, cells(5,44))
		
	elseif (cells(5,45)>0 and cells(5,48)<8) and (cells(5,45)<>cells(3,44)) Then '블록 꺼내기
		Call Eraseshadow
		Call Eraseblock
		range(cells(4+8,23), cells(9+8,28)).interior.colorindex=xlnone
		
		'보관 블록 꺼내기
		cells(3,44)=cells(5,45)
		cells(5,45)=-1
		
		'보관 블록 등장시키기
		cells(3,42)=2
		cells(3,43)=10
		cells(3,45)=1
	End If
End Sub

Sub goleft
	application.screenupdating=false
	
	If checkleft Then
		Call Eraseshadow
		Call Eraseblock
		
		cells(3,43)=cells(3,43)-1
		
		Call drawshadow
		Call drawblock
	End If
End Sub

Sub goright
	application.screenupdating=false
	
	If checkright Then
		Call Eraseshadow
		Call Eraseblock
		
		cells(3,43)=cells(3,43)+1
		
		Call drawshadow
		Call drawblock
	End If
End Sub

Sub godown
	application.screenupdating=false
	
	If checkdown Then
		Call Eraseblock
		
		cells(3,42)=cells(3,42)+1
		
		Call drawblock
	End If
End Sub

Sub gobottom
	'하드드롭 함수
	application.screenupdating=false
	
	Call Eraseshadow
	Call Eraseblock
	
	cells(3,42)=cells(3,46)
	'그림자 위치로 옮기는 것과 같음

	Call drawblock
	Call calhigh
End Sub

Sub clockturn
	'시계방향 회전
	application.screenupdating=false
	
	If checkclockturn Then
		Call Eraseshadow
		Call Eraseblock
		
		cells(3,45)=cells(3,45)-1
		
		If (cells(3,44)=1 or cells(3,44)=4 or cells(3,44)=5) and cells(3,45)=0 Then
			cells(3,45)=2
		elseif (cells(3,44)=2 or cells(3,44)=3 or cells(3,44)=6) and cells(3,45)=0 Then
			cells(3,45)=4
		End If
		
		Call drawshadow
		Call drawblock
	End If
End Sub

Sub rclockturn
	'반시계방향 회전
	application.screenupdating=false
	
	If checkrclockturn Then
		Call Eraseshadow
		Call Eraseblock
		
		cells(3,45)=cells(3,45)+1
		
		If ((cells(3,44)=1 or cells(3,44)=4 or cells(3,44)=5) and cells(3,45)=3) or _
		 ((cells(3,44)=2 or cells(3,44)=3 or cells(3,44)=6) and cells(3,45)=5) Then
			cells(3,45)=1
		End If
		
		Call drawshadow
		Call drawblock
	End If
End Sub

Sub Eraseline
	ercnt=0
	'지운 줄의 수
	score=0
	'얻은 점수
	
	For i=min(30, cells(3,42)+2) To cells(3,48) step -1
		cnt=0
		'줄에서 채워진 블록의 수
		For j=3 To 18
			If cells(i,j).interior.colorindex<>xlnone Then
				cnt=cnt+1
			End If
		Next
		
		If cnt=16 Then
			range(cells(cells(3,48)-1,3), cells(i-1,18)).copy range(cells(cells(3,48),3))
			application.cutcopymode=false
			
			ercnt=ercnt+1
			cells(3,48)=cells(3,48)+1
			i=i+1
		End If
	Next
	
	If cells(12,33)>0 Then
		score=cells(8,33)*10*(cells(12,30)+50)
	End If
	'콤보점수 반영
	
	If ercnt=0 Then
		cells(12,33)=0
	else
		cells(12,33)=cells(12,33)+1
		cells(8,35)=cells(8,35)+ercnt
		If ercnt=1 Then
			score=score+cells(8,33)*1000
		elseif ercnt=2 Then
			score=score+cells(8,33)*3000
		elseif ercnt=3 Then
			score=score+cells(8,33)*5000
		elseif ercnt=4 Then
			score=score+cells(8,33)*8000	
		End If
		
		If cells(8,35)>9 Then
			cells(8,33)=cells(8,33)+1
			cells(8,35)=cells(8,35)-10
		End If
		'레벨업 시스템
	End If
	
	cells(10, 33) = cells(10, 33) + score
    If cells(14, 33) < cells(10, 33) Then
        cells(14, 33) = cells(10, 33)
        cells(14, 34) = cells(3, 33)
        ActiveWorkbook.Save
    End If
End Sub

Sub calhigh
	'최상단 행 계산
	hightmp=0
	Select Case cells(3,44)
	Case 1
		Select Case cells(3,45)
		Case 1
			hightmp=cells(3,42)
		Case 2
			hightmp=cells(3,42)-1
		End Select
	Case 2
		Select Case cells(3,45)
		Case 1
			hightmp=cells(3,42)
		Case 2,3,4
			hightmp=cells(3,42)-1
		End Select
	Case 3
		Select Case cells(3,45)
		Case 1
			hightmp=cells(3,42)
		Case 2,3,4
			hightmp=cells(3,42)-1
		End Select
	Case 4
		Select Case cells(3,45)
		Case 1
			hightmp=cells(3,42)
		Case 2
			hightmp=cells(3,42)-1
		End Select
	Case 5
		Select Case cells(3,45)
		Case 1
			hightmp=cells(3,42)
		Case 2
			hightmp=cells(3,42)-1
		End Select
	Case 6
	Select Case cells(3,45)
		Case 1
			hightmp=cells(3,42)
		Case 2,3,4
			hightmp=cells(3,42)-1
		End Select
	Case 7
		hightmp=cells(3,42)
	End Select
	
	If cells(3,48)>hightmp Then
		cells(3,48)=hightmp
	End If
End Sub

Sub gameloop
	'메인 재귀(자동 하강) 함수
	If cells(5,42)=0 Then
		Exit Sub
	else
		application.screenupdating=false
		
		Timeterm=81-cells(8,33)
		If Timeterm<60 Then
			Timeterm=60
		End If
		
		'줄 올라온 시간이 됐을 때
		If cells(3,42)>2 and Int(cells(3,33)-cells(5,46))>=Timeterm Then
			cells(5,46)=Int(cells(3,33))
			
			Call Eraseshadow
			Call Eraseblock
			Call panaltyline
			
			If cells(3,48)>3 Then
				rtmp=cells(3,42)
				If cells(3,45)=1 Then
					cells(3,42)=3
				else
					cells(3,42)=4
				End If
				Call drawshadow
				
				If rtmp<cells(3,46) Then
					cells(3,42)=rtmp
					Call drawblock
				else
					If checkdown Then
						Call gobottom
					else
						Call drawblock
					End If
				End If
			End If
		End If
		
		If checkdown Then
			If cells(3,42)<>2 Then
				Call Eraseshadow
				Call Eraseblock
			End If
			
			cells(3,42)=cells(3,42)+1
			
			Call drawshadow
			Call drawblock
			
			cells(5,43)="=NOW()+""00:00:00.25"""
			cells(3,33)=cells(3,33)+25/85
			application.ontime cells(5,43), "gameloop"
		else
			Call calhigh
			Call Eraseline
			
			cells(3,42)=2
			cells(3,43)=10
			cells(3,44)=cells(5,44)
			cells(3,45)=1
			cells(5,44)=Int(Rnd*7)+1
			
			If cells(5,45)=-1 Then
				cells(5,45)=0
			End If
			
			range(cells(4,23), cells(9,28)).interior.colorindex=xlnone
			
			If cells(3,48)>3 and checkdown Then
				Call drawother(0,cells(5,44))
				application.screenupdating=true
				
				Call gameloop
				Exit Sub
			else
				If cells(3,48)=4 Then
					Call drawblock
				End If
				application.screenupdating=true
				
				MsgBox "GAME OVER!!!!" & vbnewline & cells(10,33) & "점을 달성하였습니다." & _
				vbnewline & "플레이 시간 : " & Int(cells(3,33)+0.5) & "초", 64, "테트리스"
				
				range(cells(2,3),cells(2,18)).interior.color=RGB(128,128,128)
				
				Call resetgame
				Exit Sub
			End If
			
		End If
		
	End If
End Sub

Sub keyset
	For i=0 To 24
		application.onkey cstr(Chr(65+i)),""
	Next
	
	For i=0 To 9
		application.onkey cstr(i),""
	Next
	
	application.onkey "{LEFT}", "goleft"
	application.onkey "{RIGHT}", "goright"
	application.onkey "{DOWN}", "godown"
	application.onkey "{UP}", "clockturn"
	application.onkey "Z", "rclockturn"
	application.onkey " ", "gobottom"
	application.onkey "X", "hold"
End Sub

Sub unkeyset
	For i=0 To 24
		application.onkey cstr(Chr(65+i))
	Next
	
	For i=0 To 9
		application.onkey cstr(i)
	Next
	
	application.onkey "{LEFT}"
	application.onkey "{RIGHT}"
	application.onkey "{DOWN}"
	application.onkey "{UP}"
	application.onkey " "
End Sub

Sub gamestart
	'게임 시작
	If cells(5,42)=0 Then
		cells(5,42)=1
		
		Call keyset
		
		For i=0 To 1
			If range(cells(4+8*i,23), cells(9+8*i,28)).interior.colorindex=xlnone Then
				application.screenupdating=false
				Call drawother(i, cells(5,44+i))
				application.screenupdating=true
			End If
		Next
		
		Call gameloop
	End If
End Sub

Sub pausegame
	'게임 일시정지
	If cells(5,42)=1 Then
		cells(5,42)=0
		
		Call unkeyset
	End If
End Sub

Sub resetgame
	cells(3,42)=2
	cells(3,43)=10
	cells(3,44)=Int(Rnd*7)+1
	cells(3,45)=1
	cells(3,48)=31
	
	range(cells(3,3), cells(30,18)).interior.colorindex=xlnone
	For i=0 To 1
		range(cells(4+8*i,23), cells(9+8*i,28)).interior.colorindex=xlnone
	Next
	
	cells(8,33)=1
	cells(8,35)=0
	cells(3,33)=0
	cells(10,33)=0
	cells(12,33)=0
	cells(3,46)=31
	cells(3,47)=1
	cells(5,44)=Int(Rnd*7)+1
	cells(5,45)=0
	cells(5,46)=0
	
	Call pausegame
End Sub

Sub drawshadow
	rtmp=cells(3,42)
	'하드드롭 알고리즘 이용
	Do While true
		If checkdown Then
			cells(3,42)=cells(3,42)+1
		else
			Exit Do
		End If
	loop
	
	Call Selectarea
	
	on error resume Next
	with Selection.interior
		.colorindex=0
		.pattern=xlpatterndown
		.patterncolor=RGB(0,0,0)
	End with
	
	cells(3,46)=cells(3,42)
	cells(3,47)=cells(3,45)
	
	cells(3,42)=rtmp
	
	cells(48,48).select
End Sub

Sub Eraseshadow
	rtmp=cells(3,42)
	ttmp=cells(3,45)
	
	cells(3,42)=cells(3,46)
	cells(3,45)=cells(3,47)
	
	Call Selectarea
	
	on error resume Next
	Selection.interior.colorindex=xlnone
	
	cells(3,42)=rtmp
	cells(3,45)=ttmp
	
	cells(48,48).select
End Sub

Sub panaltyline
	Dim flag(15)
	For i=0 To 15
		flag(i)=i+3
	Next
	
	linenum=Int(cells(8,33)/2+1)
	If linenum>6 Then
		linenum=6
	End If
	
	If cells(3,48)-linenum<=2 Then
		linenum=cells(3,48)-3
	End If
	
	If cells(3,48)<31 Then
		range(cells(cells(3,48),3), cells(30,18)).copy range(cells(cells(3,48)-linenum,3))
		application.cutcopymode=false
	End If
	
	cells(3,48)=cells(3,48)-linenum
	For a=0 To linenum-1
		For i=0 To 15
			r=Int(Rnd*2)
			tmp=flag(r)
			flag(r)=flag(i)
			flag(i)=tmp
		Next
		
		For i=3 To 18
			If emptycell(flag,i) Then
				cells(30-a,i).interior.colorindex=xlnone
			else
				cells(30-a,i).interior.color=RGB(0,0,0)
			End If
		Next
	Next
End Sub

Function emptycell(flag, row)
	emtcnt=Int(Rnd*4)
	
	For i=0 To emtcnt
		If row=flag(i) Then
			emptycell=true
			Exit Function
		End If
	Next
	emptycell=false
End Function

Sub help
	MsgBox "조작방법" & vbnewline & "X : 블록 보관하기" & vbnewline & "Z : 반시계 방향 회전" & _
	vbnewline & "위쪽 방향키 : 시계 방향 회전" & vbnewline & "아래쪽 방향키 : 낙하 속도 증가" & _
	vbnewline & "왼쪽, 오른쪽 방향키 : 좌우 이동" & vbnewline & "스페이스바 : 하드 드롭", _
	64, "테트리스"
End Sub
