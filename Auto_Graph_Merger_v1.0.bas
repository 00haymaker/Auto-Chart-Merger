Option Explicit

Sub Run_AutoChartMerger()
    Dim wsControl As Worksheet
    Dim sheetIdx As Long, groupSize As Long, startRow As Long
    Dim xCol As String, yCol As String
    Dim xTitle As String, yTitle As String, legendOpt As String
    Dim showLegend As Boolean
    
    Dim fd As FileDialog
    Dim selectedFile As Variant
    Dim totalFiles As Long, fileCount As Long, chartCount As Long
    
    Dim wbOut As Workbook, wsOut As Worksheet
    Dim wbData As Workbook, wsData As Worksheet
    Dim chObj As ChartObject
    Dim lastRow As Long
    Dim xVals As Variant, yVals As Variant
    Dim cleanName As String
    
    Dim chartLeft As Double, chartTop As Double
    Dim chartWidth As Double, chartHeight As Double
    
    Dim defaultFileName As String
    Dim savePath As Variant
    
    ' 1. 실행 시트에서 8가지 사용자 입력 변수 읽기
    Set wsControl = ThisWorkbook.Sheets(1)
    
    On Error GoTo InputError
    sheetIdx = CLng(wsControl.Range("C4").Value)
    groupSize = CLng(wsControl.Range("C5").Value)
    startRow = CLng(wsControl.Range("C6").Value)
    xCol = Trim(CStr(wsControl.Range("C7").Value))
    yCol = Trim(CStr(wsControl.Range("C8").Value))
    xTitle = Trim(CStr(wsControl.Range("C9").Value))
    yTitle = Trim(CStr(wsControl.Range("C10").Value))
    legendOpt = UCase(Trim(CStr(wsControl.Range("C11").Value)))
    On Error GoTo 0
    
    ' 기본 입력값 유효성 검사
    If sheetIdx < 1 Or groupSize < 1 Or startRow < 1 Or xCol = "" Or yCol = "" Then
        MsgBox "입력 변수(C4~C8)를 올바르게 입력해 주세요.", vbExclamation, "입력 오류"
        Exit Sub
    End If
    
    If legendOpt = "ON" Or legendOpt = "TRUE" Or legendOpt = "Y" Then
        showLegend = True
    Else
        showLegend = False
    End If
    
    ' 2. 분석 대상 파일 다중 선택 팝업창 띄우기
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "병합할 데이터 엑셀 파일들을 선택하세요 (다중 선택 가능)"
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "Excel 파일", "*.xlsx; *.xls; *.csv"
        
        If .Show = False Then Exit Sub ' 사용자가 취소한 경우 종료
        totalFiles = .SelectedItems.Count
    End With
    
    ' 성능 최적화 (화면 갱신 및 경고창 일시 중지)
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' 3. 결과 출력용 새 엑셀 파일(.xlsx용 통합 문서) 생성
    Set wbOut = Workbooks.Add
    Set wsOut = wbOut.Sheets(1)
    wsOut.Name = "Merged_Charts"
    
    ' 차트 기본 크기 설정 (가로 2열 바둑판 자동 배치용)
    chartWidth = 480
    chartHeight = 320
    fileCount = 0
    chartCount = 0
    
    ' 4. 선택된 파일 순회 및 그래프 겹치기
    For Each selectedFile In fd.SelectedItems
        
        ' 4-1. 지정된 개수 단위의 첫 번째 파일일 때마다 새 차트 생성
        If fileCount Mod groupSize = 0 Then
            ' 2열 그리드 좌표 계산 (겹침 방지)
            chartLeft = 20 + (chartCount Mod 2) * (chartWidth + 25)
            chartTop = 20 + (chartCount \ 2) * (chartHeight + 25)
            
            Set chObj = wsOut.ChartObjects.Add(Left:=chartLeft, Top:=chartTop, Width:=chartWidth, Height:=chartHeight)
            
            With chObj.Chart
                .ChartType = xlXYScatterLinesNoMarkers ' 표식 없는 직선 분산형 차트
                .HasTitle = True
                .ChartTitle.Text = "Group " & (chartCount + 1) & " (" & (fileCount + 1) & "~" & Application.Min(fileCount + groupSize, totalFiles) & ")"
                
                ' X축 제목 설정
                If xTitle <> "" Then
                    .Axes(xlCategory, xlPrimary).HasTitle = True
                    .Axes(xlCategory, xlPrimary).AxisTitle.Text = xTitle
                Else
                    .Axes(xlCategory, xlPrimary).HasTitle = False
                End If
                
                ' Y축 제목 설정
                If yTitle <> "" Then
                    .Axes(xlValue, xlPrimary).HasTitle = True
                    .Axes(xlValue, xlPrimary).AxisTitle.Text = yTitle
                Else
                    .Axes(xlValue, xlPrimary).HasTitle = False
                End If
                
                ' 범례 표시 여부 설정
                .HasLegend = showLegend
                If showLegend Then
                    .Legend.Position = xlLegendPositionRight
                End If
            End With
            
            chartCount = chartCount + 1
        End If
        
        ' 4-2. 원본 데이터 파일 읽기 전용으로 열기 및 값 추출
        Set wbData = Workbooks.Open(Filename:=CStr(selectedFile), ReadOnly:=True)
        
        If sheetIdx > wbData.Sheets.Count Then
            wbData.Close SaveChanges:=False
            GoTo CleanupWithError
        End If
        
        Set wsData = wbData.Sheets(sheetIdx)
        
        ' X열과 Y열 중 데이터가 있는 마지막 행 자동 탐색
        lastRow = Application.Max( _
            wsData.Cells(wsData.Rows.Count, xCol).End(xlUp).Row, _
            wsData.Cells(wsData.Rows.Count, yCol).End(xlUp).Row)
            
        If lastRow >= startRow Then
            xVals = wsData.Range(xCol & startRow & ":" & xCol & lastRow).Value
            yVals = wsData.Range(yCol & startRow & ":" & yCol & lastRow).Value
            
            ' 확장자를 제외한 순수 파일명 추출
            cleanName = Mid(CStr(selectedFile), InStrRev(CStr(selectedFile), "\") + 1)
            If InStrRev(cleanName, ".") > 0 Then
                cleanName = Left(cleanName, InStrRev(cleanName, ".") - 1)
            End If
            
            ' 4-3. 현재 차트에 새 곡선(계열) 추가
            With chObj.Chart.SeriesCollection.NewSeries
                .Name = cleanName
                .XValues = xVals
                .Values = yVals
            End With
        End If
        
        wbData.Close SaveChanges:=False
        fileCount = fileCount + 1
    Next selectedFile
    
    Application.ScreenUpdating = True
    
    ' 5. 결과 파일 저장 창 띄우기 (Result_날짜시간.xlsx 기본 입력)
    defaultFileName = "Result_" & Format(Now, "yyyymmdd_hhnn") & ".xlsx"
    savePath = Application.GetSaveAsFilename( _
        InitialFileName:=defaultFileName, _
        FileFilter:="Excel 통합 문서 (*.xlsx), *.xlsx", _
        Title:="병합 결과 파일 저장")
        
    If savePath <> False Then
        wbOut.SaveAs Filename:=CStr(savePath), FileFormat:=xlOpenXMLWorkbook
        MsgBox "총 " & fileCount & "개 파일 처리 및 " & chartCount & "개 그래프 생성이 완료되었습니다." & vbCrLf & _
               "저장 위치: " & CStr(savePath), vbInformation, "Auto Chart Merger 완료"
    Else
        MsgBox "저장이 취소되었습니다. 생성된 결과 창에서 직접 저장하실 수 있습니다.", vbInformation, "저장 취소"
    End If
    
    Application.DisplayAlerts = True
    Exit Sub

InputError:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "시트 번호(C4), 겹칠 개수(C5), 시작 행(C6)에는 숫자만 입력해야 합니다.", vbExclamation, "입력 형식 오류"
    Exit Sub

CleanupWithError:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "원본 파일에 " & sheetIdx & "번째 시트가 존재하지 않습니다. 시트 번호를 확인해 주세요.", vbCritical, "시트 오류"
End Sub