''''''''''''USERFOMR1'''''''''''''''''''

Private Sub CommandButton1_Click()

    Dim HeaderRow As Integer
    Dim MDMIDCol As Integer
    Dim FluidCol As Integer
    Dim SerialNoCol As Integer
    Dim SizeCol As Integer
    Dim MaterialCodeCol As Integer
    Dim InsulationCodeCol As Integer
    Dim TracingCol As Integer
    Dim JacketCodeCol As Integer

'유저폼에서 입력한 값을 변수에 저장
    HeaderRow = Me.HeaderRow.Value
    MDMIDCol = Me.MDMIDCol.Value
    FluidCol = Me.FluidCol.Value
    SerialNoCol = Me.SerialNoCol.Value
    SizeCol = Me.SizeCol.Value
    MaterialCodeCol = Me.MaterialCodeCol.Value
    InsulationCodeCol = Me.InsulationCodeCol.Value
    TracingCol = Me.TracingCol.Value
    JacketCodeCol = Me.JacketCodeCol.Value

    '유효성 검사
    If HeaderRow <= 0 Or MDMIDCol <= 0 Or FluidCol <= 0 Or SerialNoCol <= 0 Or SizeCol <= 0 Or _
        MaterialCodeCol <= 0 Or InsulationCodeCol <= 0 Or TracingCol <= 0 Or JacketCodeCol <= 0 Then
        MsgBox "입력한 값 중에 잘못된 값이 있습니다. 다시 입력해주세요.", vbExclamation, "오류"
        Exit Sub
    End If


    '프로시저 실행
    Module1.CreateMDMID HeaderRow, MDMIDCol, FluidCol, SerialNoCol, SizeCol, MaterialCodeCol, InsulationCodeCol, TracingCol, JacketCodeCol


    '유저폼 종료
    Unload Me
   
End Sub



''''''''''''''''MODULE1'''''''''''''''''''''
'sheet 등의 함수에 저장
Function IsFormLoaded(ByVal FormName As String) As Boolean
    Dim frm As Object
    For Each frm In VBA.UserForms
        If frm.Name = FormName Then
            IsFormLoaded = True
            Exit Function
        End If
    Next frm
    IsFormLoaded = False
End Function

Sub CREATMDMIDPIP()
    '사용자 정의 폼 띄우기
    If Not IsFormLoaded("UserForm1") Then
        UserForm1.Show
    End If
End Sub



Function IsFormLoaded(ByVal FormName As String) As Boolean
    Dim frm As Object
    For Each frm In VBA.UserForms
        If frm.Name = FormName Then
            IsFormLoaded = True
            Exit Function
        End If
    Next frm
    IsFormLoaded = False
End Function


Sub CreateMDMID(HeaderRow As Integer, MDMIDCol As Integer, FluidCol As Integer, SerialNoCol As Integer, SizeCol As Integer, MaterialCodeCol As Integer, InsulationCodeCol As Integer, TracingCol As Integer, JacketCodeCol As Integer)


    Dim LastRow As Long
    Dim i As Long, j As Long
    Dim MDMID As String
    Dim MDMIDs() As String
    Dim Fluid As String, SerialNo As String, Size As String, MaterialCode As String, _
        InsulationCode As String, Tracing As String, JacketCode As String
    Dim RandomNum As String
    Dim Found As Boolean


    '사용자 정의 폼 띄우기
    If Not IsFormLoaded("UserForm1") Then
        UserForm1.Show
    End If
   
    Set dictMDMIDs = CreateObject("Scripting.Dictionary")
   
    '데이터가 있는 마지막 행을 찾기
    LastRow = Cells(Rows.Count, FluidCol).End(xlUp).Row
   
    '각 행별로 MDM ID 생성
    ReDim MDMIDs(2 To LastRow) 'MDM ID 배열 크기 지정
   
    For i = HeaderRow + 1 To LastRow '2번 행부터 시작 (1번 행은 헤더)
        '각 열의 값 가져오기
        Fluid = Cells(i, FluidCol).Value
        SerialNo = Cells(i, SerialNoCol).Value
        Size = Cells(i, SizeCol).Value
        MaterialCode = Cells(i, MaterialCodeCol).Value
        InsulationCode = Cells(i, InsulationCodeCol).Value
        Tracing = Cells(i, TracingCol).Value
        JacketCode = Cells(i, JacketCodeCol).Value
   
        'MDM ID 생성
        MDMID = Join(Array(Fluid, SerialNo, Size, MaterialCode, InsulationCode, Tracing, JacketCode), "-")
        Do While InStr(MDMID, "--") > 0
            MDMID = Replace(MDMID, "--", "-")
        Loop
        If Right(MDMID, 1) = "-" Then MDMID = Left(MDMID, Len(MDMID) - 1) '마지막 "-" 제거


        '중복된 MDM ID가 있는지 체크
        If dictMDMIDs.Exists(MDMID) Then
            '이미 존재하는 MDM ID일 경우 SerialNo 랜덤 값 생성하여 MDM ID 생성
            Dim SerialLen As Long
            SerialLen = Len(SerialNo)
            Do
                RandomNum = Format(Int(Rnd * (10 ^ SerialLen)), String(SerialLen, "0"))
                MDMID = Join(Array(Fluid, RandomNum, Size, MaterialCode, InsulationCode, Tracing, JacketCode), "-")
                Do While InStr(MDMID, "--") > 0
                    MDMID = Replace(MDMID, "--", "-")
                Loop
                If Right(MDMID, 1) = "-" Then MDMID = Left(MDMID, Len(MDMID) - 1) '마지막 "-" 제거
            Loop While dictMDMIDs.Exists(MDMID)
           
            'SerialNo 값을 랜덤 값으로 변경 및 마킹
            Cells(i, SerialNoCol).Value = RandomNum
            Cells(i, SerialNoCol).Interior.Color = vbYellow
        Else
            '새로운 MDM ID일 경우 딕셔너리에 추가
            dictMDMIDs.Add MDMID, ""
        End If
       
        '생성된 MDM ID A열에 입력
        Cells(i, MDMIDCol).Value = MDMID


    Next i


    '딕셔너리 삭제
    Set dictMDMIDs = Nothing
   
    '완료 메시지 출력
    MsgBox "모든 행의 MDM ID 생성이 완료되었습니다."
End Sub
