Sub makeTagSerialNo_r1()

    Dim mdm_id_list_eis() As Variant '기계/계기/specialty'의 mdm_id 리스트
    Dim mdm_id_list_ee() As Variant  'electric equipment의 mdm_id 리스트
    Dim mdm_id_list_ep() As Variant  'electric motor의 mdm_id 리스트
    Dim no_suffix_list() As Variant  'suffix를 제외한 mdm_id 조합, 기계/계기/스페셜티 tag serial no 채번시 필요
    Dim last_row As Integer '작업 대상의 마지막 행
    Dim is_right_tagSrNo As Integer 'tag serail no 채번 변수
    Dim rw As Integer 'for 문의 작업 열
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim no_suffix_mdm_id As String 'suffix가 붙지 않은 기계/계기/스페셜티 mdm-id
    Dim mdm_id As String 'suffix가 붙어 온전한 기계/계기/스페셜티 mdm-id
    Dim f As Boolean
    Dim r As Integer
    Dim s As Integer
    Dim t As Integer
    Dim eis_counting As Integer
    Dim ee_counting As Integer
    Dim ep_counting As Integer
    Dim em_counting As Integer
    Dim ed_counting As Integer
    Dim dupl As Integer
    Dim hdr_row As Integer
    
    Dim col_mdmupload As String
    Dim col_namingrule As String
    Dim col_tagcode As String
    Dim col_tagNo As String
    Dim col_lineNo As String
    Dim col_secNo As String
    Dim col_serialNo As String
    Dim col_suffix As String
    Dim col_mdmid As String
    Dim col_tagcode_ele As String
    Dim col_pancode As String
    Dim col_serialNo_ele As String
    Dim col_suffix_ele As String
    Dim col_ele_room As String
    Dim col_voltdown As String
    Dim col_suffix_motor As String
    Dim col_connect_tagcode As String
    Dim col_connect_line As String
    Dim col_connect_sect As String
    Dim col_connect_sr As String
    Dim col_connect_suffix As String
    Dim col_connect_tagNo As String
    
    
    '해더 행 번호 입력
    hdr_row = CInt(InputBox("해더 행 번호 : "))
    
    '마지막 행 찾기
    last_row = Cells(Rows.Count, 1).End(xlUp).Row
    
    col_mdmupload = FindColLetter(hdr_row, "mdm 등록 여부")
    col_namingrule = FindColLetter(hdr_row, "Naming Rule")
    col_tagcode = FindColLetter(hdr_row, "태그 코드")
    col_tagNo = FindColLetter(hdr_row, "태그 번호")
    col_lineNo = FindColLetter(hdr_row, "태그 라인 번호")
    col_secNo = FindColLetter(hdr_row, "태그 섹션 번호")
    col_serialNo = FindColLetter(hdr_row, "태그 시리얼 번호")
    col_suffix = FindColLetter(hdr_row, "태그 접미사")
    col_mdmid = FindColLetter(hdr_row, "MDM 설비 ID")
    col_tagcode_ele = FindColLetter(hdr_row, "태그 코드(전기)")
    col_pancode = FindColLetter(hdr_row, "판넬 테크 코드")
    col_serialNo_ele = FindColLetter(hdr_row, "태그 시리얼 번호")
    col_suffix_ele = FindColLetter(hdr_row, "태그 접미사")
    col_ele_room = FindColLetter(hdr_row, "전기실 번호")
    col_voltdown = FindColLetter(hdr_row, "전기 계통 전압 강하 레벨")
    col_suffix_motor = FindColLetter(hdr_row, "태그 접미사")
    col_connect_tagcode = FindColLetter(hdr_row, "연결 설비 태그 코드")
    col_connect_line = FindColLetter(hdr_row, "연결 설비 태그 라인 번호")
    col_connect_sect = FindColLetter(hdr_row, "연결 설비 태그 섹션 번호")
    col_connect_sr = FindColLetter(hdr_row, "연결 설비 태그 시리얼 번호")
    col_connect_suffix = FindColLetter(hdr_row, "연결 설비 태그 접미사")
    col_connect_tagNo = FindColLetter(hdr_row, "부하 설비 태그 번호")

    
    '배열변수 공간 count
    
    For rw = hdr_row To last_row
        
        If CStr(Range(col_namingrule & rw).Value) = "기계/계기/Specialty" And CStr(Range(col_mdmupload & rw).Value) <> "REF" Then
            eis_counting = eis_counting + 1
        ElseIf CStr(Range(col_namingrule & rw).Value) = "E_equipment" Then
            ee_counting = ee_counting + 1
        ElseIf CStr(Range(col_namingrule & rw).Value) = "E_panel" Then
            ep_counting = ep_counting + 1
        ElseIf CStr(Range(col_namingrule & rw).Value) = "E_motor" Then
            em_counting = em_counting + 1
        ElseIf CStr(Range(col_namingrule & rw).Value) = "E_driver" Then
            ed_counting = ed_counting + 1
        Else
        End If
        
    Next rw
        
    ReDim mdm_id_list_eis(1 To eis_counting + 1)
    ReDim no_suffix_list(1 To eis_counting + 1)
    ReDim mdm_id_list_ee(1 To ee_counting + 1)
    ReDim mdm_id_list_ep(1 To ep_counting + 1)
    ReDim mdm_id_list_em(1 To em_counting + 1, 1 To 2)
    ReDim mdm_id_list_ed(1 To ed_counting + 1)
    
    
    '기존 mdm id를 배열에 추가
    r = 1 '기계/계기/스페셜티 id 배열에 mdm id 배치
    s = 1 '전기 equip id 배열에 mdm id 배치
    t = 1 '전기 panel id 배열에 mdm id 배치
    u = 1 '전기 motor 배열에 mdm id 배치
    v = 1 '전기 drive 배열에 mdm id 배치
    
    
    
    
    For rw = hdr_row + 1 To last_row
    
    If CStr(Range(col_mdmid & rw).Value) <> "" Then
    
        If CStr(Range(col_namingrule & rw).Value) = "기계/계기/Specialty" And CStr(Range(col_mdmupload & rw).Value) <> "REF" Then
        
            mdm_id_list_eis(r) = Range(col_mdmid & rw).Value
            r = r + 1
        
        ElseIf CStr(Range(col_namingrule & rw).Value) = "E_equipment" Then
            
            mdm_id_list_ee(s) = Range(col_mdmid & rw).Value
            s = s + 1
            
        ElseIf CStr(Range(col_namingrule & rw).Value) = "E_panel" Then
        
            mdm_id_list_ep(t) = Range(col_mdmid & rw).Value
            t = t + 1
            
        ElseIf CStr(Range(col_namingrule & rw).Value) = "E_motor" Then
        
            mdm_id_list_em(u, 1) = Range(col_mdmid & rw).Value
            u = u + 1
            
        ElseIf CStr(Range(col_namingrule & rw).Value) = "E_driver" Then
        
            mdm_id_list_ed(v) = Range(col_mdmid & rw).Value
            v = v + 1
        Else
        End If
    Else
    End If
    Next rw
    
    'Range("A11:CM" & last_row).Sort key1:=Cells(1, 55), order1:=xlAscending, Header:=xlYes
    'Range("A11:CM" & last_row).Sort key1:=Cells(1, 48), order1:=xlAscending, Header:=xlYes
    'Range("A11:CM" & last_row).Sort key1:=Cells(1, 41), order1:=xlAscending, Header:=xlYes
    'Range("A11:CM" & last_row).Sort key1:=Cells(1, 65), order1:=xlAscending, Header:=xlYes


    For rw = hdr_row + 1 To last_row
        'mdm 등록 대상이지만 MDM ID가 만들어지지 않은 태그에 MDM ID를 부여
        If CStr(Range(col_mdmid & rw).Value) = "" Then
            
            If CStr(Range(col_namingrule & rw).Value) = "기계/계기/Specialty" And CStr(Range(col_mdmupload & rw).Value) <> "REF" Then
         
                For i = 1 To 999
                    is_right_tagSrNo = i
                
                    no_suffix_mdm_id = Range(col_tagcode & rw) & "-" & Range(col_lineNo & rw) & Range(col_secNo & rw) & Format(is_right_tagSrNo, "000")
                    mdm_id = Range(col_tagcode & rw) & "-" & Range(col_lineNo & rw) & Range(col_secNo & rw) & Format(is_right_tagSrNo, "000") & Range(col_suffix & rw)
                
                    bool_no_suffix_list = IsInArray(no_suffix_mdm_id, no_suffix_list)
                    bool_mdm_id_list = IsInArray(mdm_id, mdm_id_list_eis)
                
                    If bool_no_suffix_list = True And bool_mdm_id_list = True Then '중복되는 mdm id일 경우
                    
                    ElseIf bool_no_suffix_list = True And bool_mdm_id_list = False Then
                
                        If IsEmpty(Range(col_suffix & rw).Value) = True Then
                    
                        Else
                            Range(col_mdmid & rw).Value = mdm_id
                            Range(col_serialNo & rw).Value = Format(is_right_tagSrNo, "000")
                            mdm_id_list_eis(r) = mdm_id '배열에 추가
                            no_suffix_list(r) = no_suffix_mdm_id '배열에 추가
                            r = r + 1
                            Exit For
                        End If

                    ElseIf bool_no_suffix_list = False And bool_mdm_id_list = False Then
                        Range(col_mdmid & rw).Value = mdm_id
                        Range(col_serialNo & rw).Value = Format(is_right_tagSrNo, "000")
                        mdm_id_list_eis(r) = mdm_id '배열에 추가
                        no_suffix_list(r) = no_suffix_mdm_id '배열에 추가
                        r = r + 1
                        Exit For
                    ElseIf bool_no_suffix_list = False And bool_mdm_id_list = True Then
                    End If
                Next i
                
            ElseIf CStr(Range(col_namingrule & rw)) = "E_equipment" Then
                
                For i = 1 To 999
                    is_right_tagSrNo = i
                    mdm_id = Range(col_tagcode_ele & rw) & "-" & "12" & Format(is_right_tagSrNo, "000")
                
                    If IsInArray(mdm_id, mdm_id_list_ee) = True Then
                    Else
                        Range(col_mdmid & rw).Value = mdm_id
                        Range(col_ele_room & rw).Value = 1
                        Range("BG" & rw).Value = 2
                        Range(col_serialNo_ele & rw).Value = Format(is_right_tagSrNo, "000")
                        mdm_id_list_ee(s) = mdm_id '배열에 추가
                        s = s + 1
                    Exit For
                    End If
                Next i
        
            ElseIf CStr(Range(col_namingrule & rw)) = "E_panel" Then
            
                For i = 1 To 999
                    is_right_tagSrNo = i
                    mdm_id = Range(col_pancode & rw) & "-" & "1" & Format(is_right_tagSrNo, "000")
                
                    If IsInArray(mdm_id, mdm_id_list_ep) = True Then
                    Else
                        Range(col_mdmid & rw).Value = mdm_id
                        Range(col_ele_room & rw).Value = 1
                        Range(col_serialNo_ele & rw).Value = Format(is_right_tagSrNo, "000")
                        mdm_id_list_ep(t) = mdm_id '배열에 추가
                        t = t + 1
                        Exit For
                    End If
                Next i

            
            ElseIf CStr(Range(col_namingrule & rw).Value) = "E_motor" Then
            
                Range(col_connect_tagcode & rw).Value = Application.WorksheetFunction.index(Range(col_tagcode & ":" & col_tagcode), Application.WorksheetFunction.Match(Range(col_connect_tagNo & rw), Range(col_tagNo & ":" & col_tagNo), 0))
                Range(col_connect_line & rw).Value = Application.WorksheetFunction.index(Range(col_lineNo & ":" & col_lineNo), Application.WorksheetFunction.Match(Range(col_connect_tagNo & rw), Range(col_tagNo & ":" & col_tagNo), 0)) & ""
                Range(col_connect_sect & rw).Value = Application.WorksheetFunction.index(Range(col_secNo & ":" & col_secNo), Application.WorksheetFunction.Match(Range(col_connect_tagNo & rw), Range(col_tagNo & ":" & col_tagNo), 0)) & ""
                Range(col_connect_sr & rw).Value = Application.WorksheetFunction.index(Range(col_serialNo & ":" & col_serialNo), Application.WorksheetFunction.Match(Range(col_connect_tagNo & rw), Range(col_tagNo & ":" & col_tagNo), 0))
                Range(col_connect_suffix & rw).Value = Application.WorksheetFunction.index(Range(col_suffix & ":" & col_suffix), Application.WorksheetFunction.Match(Range(col_connect_tagNo & rw), Range(col_tagNo & ":" & col_tagNo), 0)) & ""
            
            
                mdm_id = Range(col_connect_tagcode & rw) & Range(col_tagcode_ele & rw) & "-" & Range(col_connect_line & rw) & Range(col_connect_sect & rw) & Range(col_connect_sr & rw) & Range(col_connect_suffix & rw) & Range(col_suffix_motor & rw)
                Range(col_mdmid & rw).Value = mdm_id
                dupl = Application.WorksheetFunction.CountIf(Range(col_connect_tagNo & ":" & col_connect_tagNo), Range(col_connect_tagNo & rw))
                If dupl > 1 Then
                    i = 1
                    Do While i <= dupl
                        For re_rw = hdr_row + 1 To last_row
                
                            If CStr(Range(col_connect_tagNo & re_rw).Value) = CStr(Range(col_connect_tagNo & rw).Value) Then
                                If CStr(Range(col_suffix_motor & rw).Value) = "" Then
                                    Range(col_suffix_motor & re_rw).Value = Chr(64 + i)
                                    Range(col_mdmid & re_rw).Value = mdm_id & Chr(64 + i)
                                    i = i + 1
                                Else
                                    i = i + 1
                                End If
                            Else
                            End If
                        Next re_rw
                    Loop
                Else
                End If
            
            ElseIf CStr(Range(col_namingrule & rw).Value) = "E_driver" Then
                
                Range(col_mdmid & rw) = Application.WorksheetFunction.index(Range("AV:AV"), Application.WorksheetFunction.Match(Range(col_connect_tagNo & rw), Range(col_tagNo & ":" & col_tagNo), 0)) & "-" & Range(col_tagcode_ele & rw)
                mdm_id = Range(col_mdmid & rw)
                If IsInArray(mdm_id, mdm_id_list_ed) = True Then
                    Range(col_mdmid & rw).Interior.Color = RGB(255, 0, 0)
                Else
                End If
                mdm_id_list_ed(v) = mdm_id
                v = v + 1
            
            
            End If
            
        Else
        End If
    Next rw
    
    Erase mdm_id_list_eis
    Erase no_suffix_list
    Erase mdm_id_list_ee
    Erase mdm_id_list_ep
    Erase mdm_id_list_em
    Erase mdm_id_list_ed
    
    
End Sub

Public Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    Dim i
    For i = LBound(arr) To UBound(arr)
        If arr(i) = stringToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False

End Function

Function FindColLetter(hdr_row As Integer, search_value As Variant) As String

    Dim search_rng As Range
    Dim found_cell As Range
    Dim col_letter As String
    
    Set search_rng = ActiveSheet.Rows(hdr_row)
    
    Set found_cell = search_rng.Find(What:=search_value, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not found_cell Is Nothing Then
        col_letter = Replace(found_cell.Cells.Address(False, False), hdr_row & "", "")
        FindColLetter = col_letter
    Else
        FindColLetter = "Value not found."
    End If

End Function


