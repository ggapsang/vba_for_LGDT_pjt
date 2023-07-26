Function GetWorkbook(ByVal sFullName As String) As Workbook
    Dim sFile As String
    Dim wbReturn As Workbook

    sFile = Dir(sFullName)

    On Error Resume Next
    Set wbReturn = Workbooks(sFile)

    If wbReturn Is Nothing Then
        Set GetWorkbook = Nothing
    Else
        Set GetWorkbook = wbReturn
    End If

    On Error GoTo 0
End Function

Sub WhenCtrlCVisBoring_PIPING()

    Dim srcWb As Workbook
    Dim srcWs As Worksheet
    Dim tgtFile As Variant
    Dim tgtWb As Workbook
    Dim tgtWsTotal As Worksheet
    Dim tgtWs1 As Worksheet
    Dim tgtWs2 As Worksheet
    Dim tgtWs3 As Worksheet
    Dim tgtWs4 As Worksheet
    Dim tgtWs5 As Worksheet
    Dim tgtWs6 As Worksheet
    Dim tgtWs7 As Worksheet
    Dim tgtWs9 As Worksheet
    Dim lastRow As Long
    Dim is_mdm_upload As String
    
    is_mdm_upload = InputBox("MDM 등록 여부 선택 (O또는 △ 중 하나로 입력)")
    
    ' Turn off screen updating and calculation
    Application.Calculation = xlCalculationManual
    
    ' Set the source range and workbook
    Set srcWb = ThisWorkbook
    Set srcWs = srcWb.ActiveSheet


    ' Prompt the user to select the target workbook
    tgtFile = Application.GetOpenFilename(FileFilter:="Excel Files (*.xls*), *.xls*", Title:="Select the target workbook", MultiSelect:=False)
    If tgtFile = False Then
        MsgBox "No target workbook selected. Exiting the macro."
        Exit Sub
    End If

    ' Check if the target workbook is already open
    Set tgtWb = GetWorkbook(tgtFile)

    ' If the workbook is not open, open it
    If tgtWb Is Nothing Then
        Set tgtWb = Workbooks.Open(tgtFile)
    End If
    
    'set target worksheet
    Set tgtWsTotal = tgtWb.Sheets("재검토 리스트")
    Set tgtWs1 = tgtWb.Sheets("STEP1_SERIAL NO 확인")
    Set tgtWs2 = tgtWb.Sheets("STEP2_출처 확인")
    Set tgtWs3 = tgtWb.Sheets("STEP3_TAG NO 확인")
    Set tgtWs4 = tgtWb.Sheets("STEP4_중복확인")
    Set tgtWs5 = tgtWb.Sheets("STEP5_MDM 등록여부 확인")
    Set tgtWs6 = tgtWb.Sheets("STEP6_REF 확인")
    Set tgtWs7 = tgtWb.Sheets("STEP7_제외사유 확인_rev1")
    Set tgtWs9 = tgtWb.Sheets("2.0_STEP9_CCT오탈자 확인")

    ''1. validation table 기존값 정리
    tgtWsTotal.Range("A5:N500000").ClearContents
    tgtWsTotal.Range("A4:C4").ClearContents
    tgtWs1.Range("A5:E500000").ClearContents
    tgtWs1.Range("A4:C4").ClearContents
    tgtWs2.Range("A4:C500000").ClearContents
    tgtWs2.Range("A3:B3").ClearContents
    tgtWs3.Range("A4:F500000").ClearContents
    tgtWs3.Range("A3:C3").ClearContents
    tgtWs4.Range("A5:K500000").ClearContents
    tgtWs4.Range("A4:C4").ClearContents
    tgtWs5.Range("A5:D500000").ClearContents
    tgtWs5.Range("A4:C4").ClearContents
    tgtWs6.Range("A5:F500000").ClearContents
    tgtWs6.Range("A4:C4").ClearContents
    tgtWs7.Range("A7:T500000").ClearContents
    tgtWs7.Range("A6:E6").ClearContents
    tgtWs9.Range("B5:E500000").ClearContents
    tgtWs9.Range("B4:C4").ClearContents
    
    ''2. source sheet에서 복사-붙여넣기
    ' 0_total sheet (전체 시트)
    Range("U2:V2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    tgtWsTotal.Cells(4, 2).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    '1_tgtWs1 (SERIAL NO 확인)
    Range("A2").Select 'SR NO가 있는 첫번째 열
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    tgtWs1.Cells(4, 1).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    Range("U2").Select 'SR NO가 있는 두번째 열
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    tgtWs1.Cells(4, 2).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    Range("DC2").Select 'SR NO가 있는 마지막 열
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    tgtWs1.Cells(4, 3).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False

    '2_tgtWs2 (출처 확인)
    Range("A2:B2").Select 'SR NO와 출처가 있는 열
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    tgtWs2.Cells(3, 1).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False

    '3_tgtWs3 (TAG NO 확인)-필터링 중
    Range("A1").AutoFilter Field:=33, Criteria1:=is_mdm_upload, Operator:=xlOr, Criteria2:=is_mdm_upload & "(모름)" '2.0에서는 모름을 O로
    Range("A2").Select 'SR NO가 있는 열
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    tgtWs3.Cells(3, 1).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    Range("X2:Y2").Select 'TAG NO와 TAG NO 수정이 있는 열
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    tgtWs3.Cells(3, 2).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    '4_tgtWs4 (중복확인)-필터링 중
    Range("U2").Select 'SR NO가 있는 열
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    tgtWs4.Cells(4, 1).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    Range("Y2").Select 'TAG NO 수정이 있는 열
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    tgtWs4.Cells(4, 2).PasteSpecial Paste:=xlPasteValues
    tgtWs4.Cells(4, 3).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    '5_tgtWs5 (MDM 등록여부 확인)-필터링 중
    Range("U2:V2").Select 'SR NO와 대표 SR NO가 있는 열
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    tgtWs5.Cells(4, 1).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    Range("AG2").Select 'MDM 등록 여부가 있는 열
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    tgtWs5.Cells(4, 3).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    '7_tgWs7 (제외사유 확인)-필터링 해제
    If ActiveSheet.AutoFilterMode Then ActiveSheet.UsedRange.AutoFilter
    Range("U2:V2").Select 'SR NO와 대표 SR NO가 있는 열
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    tgtWs7.Cells(6, 1).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    Range("AA2").Select '카테고리가 있는 열
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    tgtWs7.Cells(6, 3).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Range("AD2:AD" & lastRow).Select '제외사유가 있는 열
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    tgtWs7.Cells(6, 4).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    Range("AG2").Select 'MDM 등록 여부가 있는 열
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    tgtWs7.Cells(6, 5).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    '9_tgtWs9 (CCT오탈자 확인)-필터링
    Range("A1").AutoFilter Field:=33, Criteria1:=is_mdm_upload, Operator:=xlOr, Criteria2:=is_mdm_upload & "(모름)"
    Range("U2").Select 'SR NO가 있는 열
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    tgtWs9.Cells(4, 2).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    Range("AE2").Select 'CCT |이 있는 열
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    tgtWs9.Cells(4, 3).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    If ActiveSheet.AutoFilterMode Then ActiveSheet.UsedRange.AutoFilter
    
    ' Turn screen updating and calculation back on
    Application.Calculation = xlCalculationAutomatic
    
    
    ''target sheet에서 작업 진행
    tgtWb.Activate
    '9_tgtWs9 (CCT오탈자 확인)-필터링
    tgtWs9.Activate
    lastRow = Cells(Rows.Count, 2).End(xlUp).Row
    Range("D4:E4").Copy
    Range("D5:E" & lastRow).PasteSpecial
    ' Clear clipboard
    Application.CutCopyMode = False
    
    '7_tgWs7 (제외사유 확인)-필터링 해제
    tgtWs7.Activate
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Range("F6:T6").Copy
    Range("F7:T" & lastRow).PasteSpecial
    ' Clear clipboard
    Application.CutCopyMode = False
    
    '5_tgtWs5 (MDM 등록여부 확인)-필터링 중
    tgtWs5.Activate
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Range("D4").Copy
    Range("D5:D" & lastRow).PasteSpecial
    ' Clear clipboard
    Application.CutCopyMode = False
    
    '4_tgtWs4 (중복확인)-필터링 중
    tgtWs4.Activate
    Range("C4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Replace What:="[*]", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Range("D4:K4").Copy
    Range("D5:K" & lastRow).PasteSpecial
    ' Clear clipboard
    Application.CutCopyMode = False
    
    '3_tgtWs3 (TAG NO 확인)-필터링 중
    tgtWs3.Activate
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Range("D3:F3").Copy
    Range("D4:F" & lastRow).PasteSpecial
    ' Clear clipboard
    Application.CutCopyMode = False
    
    '2_tgtWs2 (출처 확인)
    tgtWs2.Activate
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Range("C3").Copy
    Range("C4:C" & lastRow).PasteSpecial
    ' Clear clipboard
    Application.CutCopyMode = False
    
    '1_tgtWs1 (SERIAL NO 확인)
    tgtWs1.Activate
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Range("D4:E4").Copy
    Range("D5:E" & lastRow).PasteSpecial
    ' Clear clipboard
    Application.CutCopyMode = False
    
    ' 0_total sheet (전체 시트)
    tgtWsTotal.Activate
    lastRow = Cells(Rows.Count, 2).End(xlUp).Row
    Range("D4:L4").Copy
    Range("D5:L" & lastRow).PasteSpecial
    ' Clear clipboard
    Application.CutCopyMode = False


    ' Save the target workbook without closing it
    tgtWb.Save
    MsgBox "끝"


End Sub


Sub WhenCtrlCVisBoring_hdrow11()

    Dim srcWb As Workbook
    Dim srcWs As Worksheet
    Dim tgtFile As Variant
    Dim tgtWb As Workbook
    Dim tgtWsTotal As Worksheet
    Dim tgtWs1 As Worksheet
    Dim tgtWs2 As Worksheet
    Dim tgtWs3 As Worksheet
    Dim tgtWs4 As Worksheet
    Dim tgtWs5 As Worksheet
    Dim tgtWs6 As Worksheet
    Dim tgtWs7 As Worksheet
    Dim tgtWs9 As Worksheet
    Dim lastRow As Long
    Dim is_mdm_upload As String
    
    is_mdm_upload = InputBox("MDM 등록 여부 선택 (O또는 △ 중 하나로 입력)")
    
    ' Turn off screen updating and calculation
    Application.Calculation = xlCalculationManual
    
    ' Set the source range and workbook
    Set srcWb = ThisWorkbook
    Set srcWs = srcWb.ActiveSheet


    ' Prompt the user to select the target workbook
    tgtFile = Application.GetOpenFilename(FileFilter:="Excel Files (*.xls*), *.xls*", Title:="Select the target workbook", MultiSelect:=False)
    If tgtFile = False Then
        MsgBox "No target workbook selected. Exiting the macro."
        Exit Sub
    End If

    ' Check if the target workbook is already open
    Set tgtWb = GetWorkbook(tgtFile)

    ' If the workbook is not open, open it
    If tgtWb Is Nothing Then
        Set tgtWb = Workbooks.Open(tgtFile)
    End If
    
    'set target worksheet
    Set tgtWsTotal = tgtWb.Sheets("재검토 리스트")
    Set tgtWs1 = tgtWb.Sheets("STEP1_SERIAL NO 확인")
    Set tgtWs2 = tgtWb.Sheets("STEP2_출처 확인")
    Set tgtWs3 = tgtWb.Sheets("STEP3_TAG NO 확인")
    Set tgtWs4 = tgtWb.Sheets("STEP4_중복확인")
    Set tgtWs5 = tgtWb.Sheets("STEP5_MDM 등록여부 확인")
    Set tgtWs6 = tgtWb.Sheets("STEP6_REF 확인")
    Set tgtWs7 = tgtWb.Sheets("STEP7_제외사유 확인_rev1")
    Set tgtWs9 = tgtWb.Sheets("2.0_STEP9_CCT오탈자 확인")

    ''1. validation table 기존값 정리
    tgtWsTotal.Range("A5:N500000").ClearContents
    tgtWsTotal.Range("A4:C4").ClearContents
    tgtWs1.Range("A5:E500000").ClearContents
    tgtWs1.Range("A4:C4").ClearContents
    tgtWs2.Range("A4:C500000").ClearContents
    tgtWs2.Range("A3:B3").ClearContents
    tgtWs3.Range("A4:F500000").ClearContents
    tgtWs3.Range("A3:C3").ClearContents
    tgtWs4.Range("A5:K500000").ClearContents
    tgtWs4.Range("A4:C4").ClearContents
    tgtWs5.Range("A5:D500000").ClearContents
    tgtWs5.Range("A4:C4").ClearContents
    tgtWs6.Range("A5:F500000").ClearContents
    tgtWs6.Range("A4:C4").ClearContents
    tgtWs7.Range("A7:T500000").ClearContents
    tgtWs7.Range("A6:E6").ClearContents
    tgtWs9.Range("B5:E500000").ClearContents
    tgtWs9.Range("B4:C4").ClearContents
    
    ''2. source sheet에서 복사-붙여넣기
    ' 0_total sheet (전체 시트)
    Range("AC12:AD12").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    tgtWsTotal.Cells(4, 2).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    '1_tgtWs1 (SERIAL NO 확인)
    Range("A12").Select 'SR NO가 있는 첫번째 열
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    tgtWs1.Cells(4, 1).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    Range("AC12").Select 'SR NO가 있는 두번째 열
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    tgtWs1.Cells(4, 2).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    Range("CL12").Select 'SR NO가 있는 마지막 열
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    tgtWs1.Cells(4, 3).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False

    '2_tgtWs2 (출처 확인)
    Range("A12:B12").Select 'SR NO와 출처가 있는 열
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    tgtWs2.Cells(3, 1).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False

    '3_tgtWs3 (TAG NO 확인)-필터링 중
    Range("A11").AutoFilter Field:=41, Criteria1:=is_mdm_upload, Operator:=xlOr, Criteria2:=is_mdm_upload & "(모름)"
    Range("AC12").Select 'SR NO가 있는 열
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    tgtWs3.Cells(3, 1).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    Range("AG12").Select 'TAG NO 수정이 있는 열
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    tgtWs3.Cells(3, 2).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    '4_tgtWs4 (중복확인)-필터링 중
    Range("AC12").Select 'SR NO가 있는 열
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    tgtWs4.Cells(4, 1).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    Range("AG12").Select 'TAG NO 수정이 있는 열
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    tgtWs4.Cells(4, 2).PasteSpecial Paste:=xlPasteValues
    tgtWs4.Cells(4, 3).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    '5_tgtWs5 (MDM 등록여부 확인)-필터링 중
    Range("AC12:AD12").Select 'SR NO와 대표 SR NO가 있는 열
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    tgtWs5.Cells(4, 1).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    Range("AO12").Select 'MDM 등록 여부가 있는 열
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    tgtWs5.Cells(4, 3).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    
    '6_tgtWs6 (REF 확인)-필터링 중(계기)
    '필터해제
    If ActiveSheet.AutoFilterMode Then ActiveSheet.UsedRange.AutoFilter
    Range("A11").AutoFilter Field:=35, Criteria1:="INSTRUMENT"
    Range("AC12:AD12").Select 'SR NO와 대표 SR NO가 있는 열
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    tgtWs6.Cells(4, 1).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    Range("AO12").Select 'MDM 등록 여부가 있는 열
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    tgtWs6.Cells(4, 3).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    '7_tgWs7 (제외사유 확인)-필터링 해제
    If ActiveSheet.AutoFilterMode Then ActiveSheet.UsedRange.AutoFilter
    Range("AC12:AD12").Select 'SR NO와 대표 SR NO가 있는 열
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    tgtWs7.Cells(6, 1).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False '카테고리가 있는 열
    Range("AI12").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    tgtWs7.Cells(6, 3).PasteSpecial Paste:=xlPasteValues '제외사유가 있는 열
    Application.CutCopyMode = False
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Range("AL12:AL" & lastRow).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    tgtWs7.Cells(6, 4).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    Range("AO12").Select 'MDM 등록 여부가 있는 열
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    tgtWs7.Cells(6, 5).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    '9_tgtWs9 (CCT오탈자 확인)-필터링
    Range("A11").AutoFilter Field:=41, Criteria1:=is_mdm_upload, Operator:=xlOr, Criteria2:=is_mdm_upload & "(모름)"
    Range("AC12").Select 'CCT 열
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    tgtWs9.Cells(4, 2).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    Range("AM12").Select 'CCT |이 있는 열
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    tgtWs9.Cells(4, 3).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    If ActiveSheet.AutoFilterMode Then ActiveSheet.UsedRange.AutoFilter
    
    ' Turn screen updating and calculation back on
    Application.Calculation = xlCalculationAutomatic
    
    
    ''target sheet에서 작업 진행
    tgtWb.Activate
    '9_tgtWs9 (CCT오탈자 확인)-필터링
    tgtWs9.Activate
    lastRow = Cells(Rows.Count, 2).End(xlUp).Row
    Range("D4:E4").Copy
    Range("D5:E" & lastRow).PasteSpecial
    ' Clear clipboard
    Application.CutCopyMode = False
    
    '7_tgWs7 (제외사유 확인)-필터링 해제
    tgtWs7.Activate
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Range("F6:T6").Copy
    Range("F7:T" & lastRow).PasteSpecial
    ' Clear clipboard
    Application.CutCopyMode = False
    
    '6_tgtWs6 (REF 확인)-필터링 중(계기)
    tgtWs6.Activate
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Range("D4:F4").Copy
    Range("D5:F" & lastRow).PasteSpecial
    ' Clear clipboard
    Application.CutCopyMode = False
    
    '5_tgtWs5 (MDM 등록여부 확인)-필터링 중
    tgtWs5.Activate
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Range("D4").Copy
    Range("D5:D" & lastRow).PasteSpecial
    ' Clear clipboard
    Application.CutCopyMode = False
    
    '4_tgtWs4 (중복확인)-필터링 중
    tgtWs4.Activate
    Range("C4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Replace What:="[?]", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Range("D4:K4").Copy
    Range("D5:K" & lastRow).PasteSpecial
    ' Clear clipboard
    Application.CutCopyMode = False
    
    '3_tgtWs3 (TAG NO 확인)-필터링 중
    tgtWs3.Activate
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Range("D3:F3").Copy
    Range("D4:F" & lastRow).PasteSpecial
    ' Clear clipboard
    Application.CutCopyMode = False
    
    '2_tgtWs2 (출처 확인)
    tgtWs2.Activate
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Range("C3").Copy
    Range("C4:C" & lastRow).PasteSpecial
    ' Clear clipboard
    Application.CutCopyMode = False
    
    '1_tgtWs1 (SERIAL NO 확인)
    tgtWs1.Activate
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Range("D4:E4").Copy
    Range("D5:E" & lastRow).PasteSpecial
    ' Clear clipboard
    Application.CutCopyMode = False
    
    ' 0_total sheet (전체 시트)
    tgtWsTotal.Activate
    lastRow = Cells(Rows.Count, 2).End(xlUp).Row
    Range("D4:L4").Copy
    Range("D5:L" & lastRow).PasteSpecial
    ' Clear clipboard
    Application.CutCopyMode = False


    ' Save the target workbook without closing it
    tgtWb.Save
    MsgBox "끝"


End Sub


Sub WhenCtrlCVisBoring_hdrow1()

    Dim srcWb As Workbook
    Dim srcWs As Worksheet
    Dim tgtFile As Variant
    Dim tgtWb As Workbook
    Dim tgtWsTotal As Worksheet
    Dim tgtWs1 As Worksheet
    Dim tgtWs2 As Worksheet
    Dim tgtWs3 As Worksheet
    Dim tgtWs4 As Worksheet
    Dim tgtWs5 As Worksheet
    Dim tgtWs6 As Worksheet
    Dim tgtWs7 As Worksheet
    Dim tgtWs9 As Worksheet
    Dim lastRow As Long
    Dim is_mdm_upload As String
    
    is_mdm_upload = InputBox("MDM 등록 여부 선택 (O또는 △ 중 하나로 입력)")
    
    ' Turn off screen updating and calculation
    Application.Calculation = xlCalculationManual
    
    ' Set the source range and workbook
    Set srcWb = ThisWorkbook
    Set srcWs = srcWb.ActiveSheet


    ' Prompt the user to select the target workbook
    tgtFile = Application.GetOpenFilename(FileFilter:="Excel Files (*.xls*), *.xls*", Title:="Select the target workbook", MultiSelect:=False)
    If tgtFile = False Then
        MsgBox "No target workbook selected. Exiting the macro."
        Exit Sub
    End If

    ' Check if the target workbook is already open
    Set tgtWb = GetWorkbook(tgtFile)

    ' If the workbook is not open, open it
    If tgtWb Is Nothing Then
        Set tgtWb = Workbooks.Open(tgtFile)
    End If
    
    'set target worksheet
    Set tgtWsTotal = tgtWb.Sheets("재검토 리스트")
    Set tgtWs1 = tgtWb.Sheets("STEP1_SERIAL NO 확인")
    Set tgtWs2 = tgtWb.Sheets("STEP2_출처 확인")
    Set tgtWs3 = tgtWb.Sheets("STEP3_TAG NO 확인")
    Set tgtWs4 = tgtWb.Sheets("STEP4_중복확인")
    Set tgtWs5 = tgtWb.Sheets("STEP5_MDM 등록여부 확인")
    Set tgtWs6 = tgtWb.Sheets("STEP6_REF 확인")
    Set tgtWs7 = tgtWb.Sheets("STEP7_제외사유 확인_rev1")
    Set tgtWs9 = tgtWb.Sheets("2.0_STEP9_CCT오탈자 확인")

    ''1. validation table 기존값 정리
    tgtWsTotal.Range("A5:N500000").ClearContents
    tgtWsTotal.Range("A4:C4").ClearContents
    tgtWs1.Range("A5:E500000").ClearContents
    tgtWs1.Range("A4:C4").ClearContents
    tgtWs2.Range("A4:C500000").ClearContents
    tgtWs2.Range("A3:B3").ClearContents
    tgtWs3.Range("A4:F500000").ClearContents
    tgtWs3.Range("A3:C3").ClearContents
    tgtWs4.Range("A5:K500000").ClearContents
    tgtWs4.Range("A4:C4").ClearContents
    tgtWs5.Range("A5:D500000").ClearContents
    tgtWs5.Range("A4:C4").ClearContents
    tgtWs6.Range("A5:F500000").ClearContents
    tgtWs6.Range("A4:C4").ClearContents
    tgtWs7.Range("A7:T500000").ClearContents
    tgtWs7.Range("A6:E6").ClearContents
    tgtWs9.Range("B5:E500000").ClearContents
    tgtWs9.Range("B4:C4").ClearContents
    
    ''2. source sheet에서 복사-붙여넣기
    ' 0_total sheet (전체 시트)
    Range("AC2:AD2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    tgtWsTotal.Cells(4, 2).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    '1_tgtWs1 (SERIAL NO 확인)
    Range("A2").Select 'SR NO가 있는 첫번째 열
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    tgtWs1.Cells(4, 1).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    Range("AC2").Select 'SR NO가 있는 두번째 열
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    tgtWs1.Cells(4, 2).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    Range("CL2").Select 'SR NO가 있는 마지막 열
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    tgtWs1.Cells(4, 3).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False

    '2_tgtWs2 (출처 확인)
    Range("A2:B2").Select 'SR NO와 출처가 있는 열
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    tgtWs2.Cells(3, 1).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False

    '3_tgtWs3 (TAG NO 확인)-필터링 중
    Range("A1").AutoFilter Field:=41, Criteria1:=is_mdm_upload, Operator:=xlOr, Criteria2:=is_mdm_upload & "(모름)"
    Range("AC2").Select 'SR NO가 있는 열
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    tgtWs3.Cells(3, 1).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    Range("AF2").Select 'TAG NO 수정이 있는 열
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    tgtWs3.Cells(3, 2).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    '4_tgtWs4 (중복확인)-필터링 중
    Range("AC2").Select 'SR NO가 있는 열
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    tgtWs4.Cells(4, 1).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    Range("AG2").Select 'TAG NO 수정이 있는 열
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    tgtWs4.Cells(4, 2).PasteSpecial Paste:=xlPasteValues
    tgtWs4.Cells(4, 3).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    '5_tgtWs5 (MDM 등록여부 확인)-필터링 중
    Range("AC2:AD2").Select 'SR NO와 대표 SR NO가 있는 열
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    tgtWs5.Cells(4, 1).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    Range("AO2").Select 'MDM 등록 여부가 있는 열
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    tgtWs5.Cells(4, 3).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    
    '6_tgtWs6 (REF 확인)-필터링 중(계기)
    '필터해제
    If ActiveSheet.AutoFilterMode Then ActiveSheet.UsedRange.AutoFilter
    Range("A1").AutoFilter Field:=35, Criteria1:="INSTRUMENT"
    Range("AC2:AD2").Select 'SR NO와 대표 SR NO가 있는 열
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    tgtWs6.Cells(4, 1).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    Range("AO2").Select 'MDM 등록 여부가 있는 열
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    tgtWs6.Cells(4, 3).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    '7_tgWs7 (제외사유 확인)-필터링 해제
    If ActiveSheet.AutoFilterMode Then ActiveSheet.UsedRange.AutoFilter
    Range("AC2:AD2").Select 'SR NO와 대표 SR NO가 있는 열
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    tgtWs7.Cells(6, 1).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False '카테고리가 있는 열
    Range("AI2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    tgtWs7.Cells(6, 3).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Range("AL2:AL" & lastRow).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    tgtWs7.Cells(6, 4).PasteSpecial Paste:=xlPasteValues '제외사유가 있는 열
    Application.CutCopyMode = False
    Range("AO2").Select 'MDM 등록 여부가 있는 열
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    tgtWs7.Cells(6, 5).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    '9_tgtWs9 (CCT오탈자 확인)-필터링
    Range("A1").AutoFilter Field:=41, Criteria1:=is_mdm_upload, Operator:=xlOr, Criteria2:=is_mdm_upload & "(모름)"
    Range("AC2").Select 'CCT열
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    tgtWs9.Cells(4, 2).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    Range("AM2").Select 'CCT |이 있는 열
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    tgtWs9.Cells(4, 3).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    If ActiveSheet.AutoFilterMode Then ActiveSheet.UsedRange.AutoFilter
    
    
    ' Turn screen updating and calculation back on
    Application.Calculation = xlCalculationAutomatic

    ' Save the target workbook without closing it
    tgtWb.Save

    ''target sheet에서 작업 진행
    tgtWb.Activate
    '9_tgtWs9 (CCT오탈자 확인)-필터링
    tgtWs9.Activate
    lastRow = Cells(Rows.Count, 2).End(xlUp).Row
    Range("D4:E4").Copy
    Range("D5:E" & lastRow).PasteSpecial
    ' Clear clipboard
    Application.CutCopyMode = False
    
    '7_tgWs7 (제외사유 확인)-필터링 해제
    tgtWs7.Activate
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Range("F6:T6").Copy
    Range("F7:T" & lastRow).PasteSpecial
    ' Clear clipboard
    Application.CutCopyMode = False
    
    '6_tgtWs6 (REF 확인)-필터링 중(계기)
    tgtWs6.Activate
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Range("D4:F4").Copy
    Range("D5:F" & lastRow).PasteSpecial
    ' Clear clipboard
    Application.CutCopyMode = False
    
    '5_tgtWs5 (MDM 등록여부 확인)-필터링 중
    tgtWs5.Activate
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Range("D4").Copy
    Range("D5:D" & lastRow).PasteSpecial
    ' Clear clipboard
    Application.CutCopyMode = False
    
    '4_tgtWs4 (중복확인)-필터링 중
    tgtWs4.Activate
    Range("C4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Replace What:="[?]", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Range("D4:K4").Copy
    Range("D5:K" & lastRow).PasteSpecial
    ' Clear clipboard
    Application.CutCopyMode = False
    
    '3_tgtWs3 (TAG NO 확인)-필터링 중
    tgtWs3.Activate
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Range("D3:F3").Copy
    Range("D4:F" & lastRow).PasteSpecial
    ' Clear clipboard
    Application.CutCopyMode = False
    
    '2_tgtWs2 (출처 확인)
    tgtWs2.Activate
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Range("C3").Copy
    Range("C4:C" & lastRow).PasteSpecial
    ' Clear clipboard
    Application.CutCopyMode = False
    
    '1_tgtWs1 (SERIAL NO 확인)
    tgtWs1.Activate
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Range("D4:E4").Copy
    Range("D5:E" & lastRow).PasteSpecial
    ' Clear clipboard
    Application.CutCopyMode = False
    
    ' 0_total sheet (전체 시트)
    tgtWsTotal.Activate
    lastRow = Cells(Rows.Count, 2).End(xlUp).Row
    Range("D4:L4").Copy
    Range("D5:L" & lastRow).PasteSpecial
    ' Clear clipboard
    Application.CutCopyMode = False

    ' Save the target workbook without closing it
    tgtWb.Save
    MsgBox "끝"





End Sub








