Sub InsertDataFromExcelToPowerPointSlides()
    Dim pptApp As Object
    Dim pptPresentation As Object
    Dim pptTable As Object
    Dim excelApp As Object
    Dim excelWorkbook As Object
    Dim excelWorksheet As Object
    Dim startSlideIndex As Long
    Dim endSlideIndex As Long
    Dim currentSlideIndex As Long
    Dim pptRowIndex As Long ' ppt 테이블의 행 인덱스
    Dim colIndex As Long
    Dim maxRows As Long
    Dim maxCols As Long

    ' 시작 슬라이드와 끝 슬라이드 지정
    startSlideIndex = 1 ' 시작 슬라이드 인덱스
    endSlideIndex = 3 ' 끝 슬라이드 인덱스

    ' 엑셀 데이터 행 인덱스 초기화
    excelRowIndex = 1 ' 엑셀 데이터의 첫 번째 행부터 시작

    ' PowerPoint 프레젠테이션 열기
    Set pptApp = GetObject(, "PowerPoint.Application")
    Set pptPresentation = pptApp.ActivePresentation

    ' 사용할 엑셀 파일 열기
    Set excelApp = CreateObject("Excel.Application")
    Set excelWorkbook = excelApp.Workbooks.Open("C:\Users\jslee97\Desktop\통합 문서2.xlsx") ' 엑셀 파일 경로 및 파일명을 수정하세요

    ' 엑셀 워크시트 선택 (시트 이름을 수정하세요)
    Set excelWorksheet = excelWorkbook.Sheets("Sheet1")

    ' 시작 슬라이드부터 끝 슬라이드까지 데이터 입력
    For currentSlideIndex = startSlideIndex To endSlideIndex
        ' 슬라이드의 표에 따라 최대 행 및 열 수 설정
        Set pptTable = pptPresentation.Slides(currentSlideIndex).Shapes(2).Table ' 현재 슬라이드의 표 인덱스 수정
        maxRows = pptTable.Rows.Count
        maxCols = pptTable.Columns.Count
        
        pptRowIndex = 2 ' ppt 테이블의 두 번째 행부터 시작

        ' 데이터 채우기
        While pptRowIndex <= maxRows
            For colIndex = 1 To maxCols
                pptTable.Cell(pptRowIndex, colIndex).Shape.TextFrame.TextRange.Text = excelWorksheet.Cells(excelRowIndex, colIndex).Value
            Next colIndex
            pptRowIndex = pptRowIndex + 1
            excelRowIndex = excelRowIndex + 1 ' 다음 엑셀 데이터 행으로 이동
        Wend
    Next currentSlideIndex

    ' 리소스 정리
    excelWorkbook.Close SaveChanges:=False
    Set excelWorksheet = Nothing
    Set excelWorkbook = Nothing
    excelApp.Quit
    Set excelApp = Nothing
    Set pptTable = Nothing
    Set pptPresentation = Nothing
    Set pptApp = Nothing

    MsgBox "데이터가 PowerPoint 슬라이드에 복사되었습니다."
End Sub

