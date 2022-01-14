Attribute VB_Name = "Module1"
    'BO_Adder by Diego Espitia, January 2022
    Global work_Sheet As Worksheet
    Global work_Book As ThisWorkbook
    Global incorrect_Account_Length As Boolean
    Public final_Row_Data
    Option Compare Text
Sub Main() 'Main BO_Adder
    Dim cell_Data As Range
    Dim data_Range As Range
    incorrect_Acount_Length = False

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    'Call check_Columns_For_Text
    Create_Final_Sheet
    Check_CID
    Account_Number
    
    If incorrect_Account_Length Then
        Exit Sub
    End If
    
    Set work_Sheet = Worksheets("Sheet1")
    Set final_Sheet = Worksheets("Final_Sheet")
    
    final_Row_Data = work_Sheet.Cells(Rows.Count, "T").End(xlUp).Row
    Set column_A_To_T = final_Sheet.Range(final_Sheet.Cells(2, "A"), final_Sheet.Cells(final_Row_Data, "T"))

    column_A_To_T.Replace What:="`", Replacement:="", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

    column_A_To_T.Replace What:="!", Replacement:="", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

    column_A_To_T.Replace What:="@", Replacement:="AT", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

    column_A_To_T.Replace What:="#", Replacement:="", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

    column_A_To_T.Replace What:="$", Replacement:="", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

    column_A_To_T.Replace What:="%", Replacement:="", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

    column_A_To_T.Replace What:="^", Replacement:="", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

    column_A_To_T.Replace What:="&", Replacement:="AND", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

    column_A_To_T.Replace What:="  ", Replacement:=" ", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

    column_A_To_T.Replace What:="Ä", Replacement:="A", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

    column_A_To_T.Replace What:="Ê", Replacement:="E", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

    column_A_To_T.Replace What:="Ï", Replacement:="I", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

    column_A_To_T.Replace What:="Ö", Replacement:="O", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

    column_A_To_T.Replace What:="Ü", Replacement:="U", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

    column_A_To_T.Replace What:="Ÿ", Replacement:="Y", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

    With column_A_To_T.Cells
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    work_Sheet.Rows("1:1").Copy
    final_Sheet.Rows("1:1").PasteSpecial xlPasteColumnWidths
    final_Sheet.Rows("1:1").PasteSpecial xlPasteValues
    final_Sheet.Rows("1:1").PasteSpecial xlFormats
    With final_Sheet.Range(final_Sheet.Cells(2, "A"), final_Sheet.Cells(final_Row_Data, "T")).Borders
        .LineStyle = xlContinuous
        .Color = black
        .Weight = xlThin
    End With
    Worksheets("Final_Sheet").Select
    Application.ScreenUpdating = True
End Sub
Private Sub check_Columns_For_Text() 'Check if columns are empty
    Dim work_Book As ThisWorkbook
    Dim work_Sheet As Worksheet
    Dim calculate_Sheet As Worksheet
    Dim final_Sheet As Worksheet
    Set work_Book = ThisWorkbook
    Set work_Sheet = work_Book.Worksheets("Sheet1")
    Set calculate_Sheet = work_Book.Worksheets("Calculate_Sheet")
    Set final_Sheet = work_Book.Worksheets("Final_Sheet")
    Dim dRng As Range, lRow As Long
    Dim kRng As Range
    Dim Col As Long
    Dim empty_Column As Boolean
    
    lRow = Cells(Rows.Count, 1).End(xlUp).Row
    Col = Cells(3, Columns.Count).End(xlToLeft).Column
    On Error Resume Next
        Set dRng = Range("A3:A" & lRow).SpecialCells(xlBlanks)
        Set kRng = Range("3:AA" & Col).SpecialCells(xlBlanks)
    On Error GoTo 0
    For Counter = 1 To Col
        For Counter_B = 3 To lRow
            If work_Sheet.Cells(Counter_B, Counter).Value <> "" Then
                empty_Column = False
                'work_Sheet.Cells(Counter_B, Counter).Interior.ColorIndex = 5
                'MsgBox ("Row " & Counter_B & " Column " & Counter)
            End If
        Next Counter_B
        'MsgBox (Counter)
        'MsgBox "Column " & Counter & " Empty: " & IsEmpty(Range("D3:D"))
        If empty_Column = True Then
            'MsgBox ("Column " & Counter & " is Empty")
            ''MsgBox (Counter_B)
        End If
        empty_Column = True
    Next Counter
    
    If Not dRng Is Nothing Then
        MsgBox ("Skipping Row")
        Exit Sub
    End If
    If Not kRng Is Nothing Then
        MsgBox ("Skipping Column")
        Exit Sub
    End If
End Sub
Private Sub Account_Number()
    Dim range_Client_ID
    Dim work_Book As Workbook
    Dim work_Sheet As Worksheet
    Dim final_Sheet As Worksheet
    Set work_Book = ThisWorkbook
    Set work_Sheet = work_Book.Worksheets("Sheet1")
    Set final_Sheet = work_Book.Worksheets("Final_Sheet")
    final_Row_Data = work_Sheet.Cells(Rows.Count, "B").End(xlUp).Row
    
    If final_Row_Data <> 1 Then
        work_Sheet.Range(work_Sheet.Cells(final_Row_Data, "B"), work_Sheet.Cells(2, "B")).Copy
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "B"), final_Sheet.Cells(2, "B")).PasteSpecial xlPasteValues
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "B"), final_Sheet.Cells(2, "B")).PasteSpecial xlPasteFormats
        
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "B"), final_Sheet.Cells(2, "B")).Replace What:="`,!,@,#,$,%,^", Replacement:="", LookAt:= _
            xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
            
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "B"), final_Sheet.Cells(2, "B")).Replace What:="@", Replacement:="AT", LookAt:= _
            xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
            
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "B"), final_Sheet.Cells(2, "B")).Replace What:="&", Replacement:="AND", LookAt:= _
            xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
            
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "B"), final_Sheet.Cells(2, "B")).Replace What:="  ", Replacement:=" ", LookAt:= _
            xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
            
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "C"), final_Sheet.Cells(2, "C")).Value = _
            "=IF(OR(RIGHT(TRIM(RC[-1]),1)="")"",RIGHT(TRIM(RC[-1]),1)=""."",RIGHT(TRIM(RC[-1]),1)="",""),LEFT(TRIM(RC[-1]),LEN(TRIM(RC[-1]))-1),TRIM(RC[-1]))"
        
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "C"), final_Sheet.Cells(2, "C")).Copy
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "B"), final_Sheet.Cells(2, "B")).PasteSpecial xlPasteValues
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "C"), final_Sheet.Cells(2, "C")).ClearContents
        
        Account_Length
    End If
End Sub
Private Sub Check_CID() 'Checking for CID
    Dim work_Book As Workbook
    Dim work_Sheet As Worksheet
    Dim final_Sheet As Worksheet
    Set work_Book = ThisWorkbook
    Set work_Sheet = work_Book.Worksheets("Sheet1")
    Set final_Sheet = work_Book.Worksheets("Final_Sheet")
    final_Row_Data = work_Sheet.Cells(Rows.Count, 1).End(xlUp).Row
    
    work_Sheet.Range(work_Sheet.Cells(final_Row_Data, "A"), work_Sheet.Cells(2, "A")).Copy
    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "A"), final_Sheet.Cells(2, "A")).PasteSpecial xlPasteValues
    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "A"), final_Sheet.Cells(2, "A")).PasteSpecial xlPasteFormats
        
    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "B"), final_Sheet.Cells(2, "B")).Value = _
        "=IF(OR(RIGHT(TRIM(RC[-1]),1)="")"",RIGHT(TRIM(RC[-1]),1)=""."",RIGHT(TRIM(RC[-1]),1)="",""),LEFT(TRIM(RC[-1]),LEN(TRIM(RC[-1]))-1),TRIM(RC[-1]))"
    
    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "B"), final_Sheet.Cells(2, "B")).Copy
    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "A"), final_Sheet.Cells(2, "A")).PasteSpecial xlPasteValues
    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "B"), final_Sheet.Cells(2, "B")).ClearContents
    
End Sub
Private Sub Account_Length() 'Check if Account Number Length is correct for User_Input CID
    Dim account_Num_Range As Range
    Dim account_Num_Cell As Range
    Dim work_Book As Workbook
    Dim work_Sheet As Worksheet
    Dim final_Sheet As Worksheet
    Set work_Book = ThisWorkbook
    Set work_Sheet = work_Book.Worksheets("Sheet1")
    Set final_Sheet = work_Book.Worksheets("Final_Sheet")
    Dim Account_Num_Length As String
    Dim CID As String
    CID = InputBox("Enter CID")
    final_Row_Data = work_Sheet.Cells(work_Sheet.Rows.Count, "B").End(xlUp).Row
    Set account_Num_Range = work_Sheet.Range(work_Sheet.Cells(2, 2), work_Sheet.Cells(final_Row_Data, 2))
    
    For Each account_Num_Cell In account_Num_Range.Cells
        If Len(account_Num_Cell.Value) <> 5 And CID = "55P" Then '---------------------------------------------------------55P
            incorrect_Account_Length = True
            work_Sheet.Cells(account_Num_Cell.Address, "B").Interior.ColorIndex = 35
        End If
        If Len(account_Num_Cell.Value) <> 11 And CID = "5DU" Then '---------------------------------------------------------5DU
            incorrect_Account_Length = True
            work_Sheet.Cells(account_Num_Cell.Row, "B").Interior.ColorIndex = 35
        End If
        If Len(account_Num_Cell.Value) <> 16 And CID = "11Z" Then '---------------------------------------------------------11Z
            incorrect_Account_Length = True
            work_Sheet.Cells(account_Num_Cell.Row, "B").Interior.ColorIndex = 35
        End If
    Next account_Num_Cell
    
    If incorrect_Account_Length Then
        MsgBox ("Cells with Incorrect Account Length have been hightlighted and sorted for you")
        work_Sheet.Range("B:B").AutoFilter Field:=1, Criteria1:=RGB(204, 255 _
        , 204), Operator:=xlFilterCellColor
    End If
End Sub
Private Sub Create_Final_Sheet() 'Create a Final_Sheet if none exist
    Dim work_Book As Workbook
    Set work_Book = ThisWorkbook
    Dim sheet_Calculate As String
    Dim sheet_Final As String
    sheet_Final = "Final_Sheet"
    Dim sheet_Final_Exists As Boolean

    For Each work_Sheet In work_Book.Worksheets
        If work_Sheet.Name = "Final_Sheet" Then
            sheet_Final_Exists = True
        End If
    Next work_Sheet
    
    If sheet_Final_Exists = False Then
        work_Book.Worksheets.Add(After:=work_Book.Worksheets("Sheet1")).Name = "Final_Sheet"
    End If
End Sub
