Attribute VB_Name = "Module1"
Global work_Book
Global work_Sheet
Sub Main()
    Dim Client_ID
    Client_ID = InputBox("Enter CID")

    Application.ScreenUpdating = False
    Call Add_Calculation_Sheet
    Call CID_Length
    Call Check_CID
    Call Account_Number
    Call Account_Name
    Call Interal_Account_Number
    Call Check_PEID
    Call Account_Flags
    'Call Account_Length
    
    Worksheets("Sheet1").Rows("1:1").Copy
    Worksheets("Final_Sheet").Rows("1:1").PasteSpecial xlPasteColumnWidths
    Worksheets("Final_Sheet").Rows("1:1").PasteSpecial xlPasteValues
    Worksheets("Final_Sheet").Rows("1:1").PasteSpecial xlFormats
    Application.ScreenUpdating = False
End Sub
Private Sub Add_Calculation_Sheet()
    Set work_Book = ThisWorkbook
    Dim sheet_Calculate As String
    Dim sheet_Final As String
    sheet_Calculate = "Calculate_Sheet"
    sheet_Final = "Final_Sheet"
    Dim sheet_Calculate_Exists As Boolean
    Dim sheet_Final_Exists As Boolean

    For Each work_Sheet In work_Book.Worksheets
        If work_Sheet.Name = "Calculate_Sheet" Then
            sheet_Calculate_Exists = True
        End If
        If work_Sheet.Name = "Final_Sheet" Then
            sheet_Final_Exists = True
        End If
    Next work_Sheet
    
    If sheet_Calculate_Exists = False Then
        work_Book.Worksheets.Add After:=Worksheets("Sheet1")
        work_Book.Worksheets(Sheets.Count).Name = "Calculate_Sheet"
        'ActiveSheet.Name = "Calculate_Sheet"
        'Worksheets("Sheet1").Select
    End If
    
    If sheet_Final_Exists = False Then
        work_Book.Worksheets.Add After:=Worksheets("Calculate_Sheet")
        work_Book.Worksheets(Sheets.Count).Name = "Final_Sheet"
        'ActiveSheet.Name = "Final_Sheet"
        'Worksheets("Sheet1").Select
    End If
End Sub
Private Sub Check_CID() 'Checking for CID
    Dim Final_Row_Client_ID
    Dim work_Book As Workbook
    Dim work_Sheet As Worksheet
    Dim calculate_Sheet As Worksheet
    Dim final_Sheet As Worksheet
    Set work_Book = ActiveWorkbook
    Set work_Sheet = work_Book.Worksheets("Sheet1")
    Set calculate_Sheet = work_Book.Worksheets("Calculate_Sheet")
    Set final_Sheet = work_Book.Worksheets("Final_Sheet")
    Final_Row_Client_ID = work_Sheet.Cells(Rows.Count, "A").End(xlUp).Row
    
    For Counter = 2 To Final_Row_Client_ID
        calculate_Sheet.Cells(Counter, 1).Value = work_Sheet.Cells(Counter, "A").Value
        calculate_Sheet.Cells(Counter, 2).Value = calculate_Sheet.Cells(Counter, 1).Value
        
        calculate_Sheet.Cells(Counter, 2).Replace What:="`", Replacement:="", LookAt:=xlPart, SearchOrder:= _
            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
            
        calculate_Sheet.Cells(Counter, 2).Replace What:="!", Replacement:="", LookAt:=xlPart, SearchOrder:= _
            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
            
        calculate_Sheet.Cells(Counter, 2).Replace What:="@", Replacement:="AT", LookAt:=xlPart, SearchOrder:= _
            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
            
        calculate_Sheet.Cells(Counter, 2).Replace What:="#", Replacement:="", LookAt:=xlPart, SearchOrder:= _
            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
            
        calculate_Sheet.Cells(Counter, 2).Replace What:="$", Replacement:="", LookAt:=xlPart, SearchOrder:= _
            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
            
        calculate_Sheet.Cells(Counter, 2).Replace What:="%", Replacement:="", LookAt:=xlPart, SearchOrder:= _
            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
            
        calculate_Sheet.Cells(Counter, 2).Replace What:="^", Replacement:="", LookAt:=xlPart, SearchOrder:= _
            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
            
        calculate_Sheet.Cells(Counter, 2).Replace What:="&", Replacement:="AND", LookAt:=xlPart, SearchOrder:= _
            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
            
        calculate_Sheet.Cells(Counter, 2).Replace What:="  ", Replacement:=" ", LookAt:=xlPart, SearchOrder:= _
            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
            
        calculate_Sheet.Cells(Counter, 3).Value = "=IF(OR(RIGHT(TRIM(RC[-1]),1)="")"",RIGHT(TRIM(RC[-1]),1)=""."",RIGHT(TRIM(RC[-1]),1)="",""),LEFT(TRIM(RC[-1]),LEN(TRIM(RC[-1]))-1),TRIM(RC[-1]))"
                
        calculate_Sheet.Cells(Counter, 3).Copy
        calculate_Sheet.Cells(Counter, 4).PasteSpecial xlPasteValues
            
        With calculate_Sheet.Cells(Counter, 4)
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
        'Final Sheet
        final_Sheet.Cells(Counter, "A").Value = Worksheets("Calculate_Sheet").Cells(Counter, 4).Value
    Next Counter
End Sub
Private Sub CID_Length()
    Dim Final_Row_CID_Length As Long
    Dim work_Book As Workbook
    Dim work_Sheet As Worksheet
    Set work_Book = ActiveWorkbook
    Set work_Sheet = work_Book.Worksheets("Sheet1")
    'Dim Account_Num_Length As String
    'Dim CID As String
    Dim Incorrect_Account_Length As Boolean
    'CID = InputBox("Enter CID")
    Final_Row_CID_Length = work_Sheet.Cells(Rows.Count, 1).End(xlUp).Row
    
    For Counter = 2 To Final_Row_CID_Length
        If Len(work_Sheet.Cells(Counter, 1).Value) <> 3 Then
            Incorrect_CID_Length = True
            work_Sheet.Cells(Counter, 1).Interior.ColorIndex = 35
        End If
        'If Len(Cells(Counter, 2).Value) <> 5 And CID = "55P" Then
        '    Incorrect_Account_Length = True
        '    Worksheets("Sheet1").Cells(Counter, 2).Interior.ColorIndex = 35
        'End If
        'If Len(Cells(Counter, 2).Value) <> 11 And CID = "5DU" Then
        '    Incorrect_Account_Length = True
        '    Worksheets("Sheet1").Cells(Counter, 2).Interior.ColorIndex = 35
        'End If
    Next Counter
    If Incorrect_CID_Length Then
        MsgBox ("Cells with Incorrect CID Length have been hightlighted and sorted for you")
        Incorrect_Account_Length = False
        work_Sheet.Range("A:A").AutoFilter Field:=1, Criteria1:=RGB(204, 255 _
        , 204), Operator:=xlFilterCellColor
    End If
End Sub
Private Sub Account_Number()
    Dim Final_Row_Account_Number
    Dim work_Book As Workbook
    Dim work_Sheet As Worksheet
    Dim calculate_Sheet As Worksheet
    Dim final_Sheet As Worksheet
    Set work_Book = ActiveWorkbook
    Set work_Sheet = work_Book.Worksheets("Sheet1")
    Set calculate_Sheet = work_Book.Worksheets("Calculate_Sheet")
    Set final_Sheet = work_Book.Worksheets("Final_Sheet")
    Dim Incorrect_Account_Length As Boolean
    Final_Row_Account_Number = work_Sheet.Cells(Rows.Count, "B").End(xlUp).Row
    
    For Counter = 2 To Final_Row_Account_Number
        calculate_Sheet.Cells(Counter, 1).Value = work_Sheet.Cells(Counter, "B").Value
        calculate_Sheet.Cells(Counter, 2).Value = Worksheets("Calculate_Sheet").Cells(Counter, 1).Value

        If Len(work_Sheet.Cells(Counter, "B").Value) <> 6 And Client_ID <> "55P" Then
            Incorrect_Account_Length = True
            work_Sheet.Cells(Counter, "B").Interior.ColorIndex = 35
        End If
'        Worksheets("Calculate_Sheet").Cells(Counter, 2).Replace What:="`", Replacement:="", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        Worksheets("Calculate_Sheet").Cells(Counter, 2).Replace What:="!", Replacement:="", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        Worksheets("Calculate_Sheet").Cells(Counter, 2).Replace What:="@", Replacement:="AT", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        Worksheets("Calculate_Sheet").Cells(Counter, 2).Replace What:="#", Replacement:="", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        Worksheets("Calculate_Sheet").Cells(Counter, 2).Replace What:="$", Replacement:="", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        Worksheets("Calculate_Sheet").Cells(Counter, 2).Replace What:="%", Replacement:="", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        Worksheets("Calculate_Sheet").Cells(Counter, 2).Replace What:="^", Replacement:="", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        Worksheets("Calculate_Sheet").Cells(Counter, 2).Replace What:="&", Replacement:="AND", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
        calculate_Sheet.Cells(Counter, 2).Replace What:="  ", Replacement:=" ", LookAt:=xlPart, SearchOrder:= _
            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

        calculate_Sheet.Cells(Counter, 3).Value = "=IF(OR(RIGHT(TRIM(RC[-1]),1)="")"",RIGHT(TRIM(RC[-1]),1)=""."",RIGHT(TRIM(RC[-1]),1)="",""),LEFT(TRIM(RC[-1]),LEN(TRIM(RC[-1]))-1),TRIM(RC[-1]))"

        calculate_Sheet.Cells(Counter, 3).Copy
        calculate_Sheet.Cells(Counter, 4).PasteSpecial xlPasteValues
            
        With calculate_Sheet.Cells(Counter, 4)
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
        'Final Sheet
        calculate_Sheet.Cells(Counter, 4).Copy
        final_Sheet.Cells(Counter, "B").PasteSpecial xlPasteValues
        final_Sheet.Cells(Counter, "B").PasteSpecial xlFormats
        'Filter Incorrect Accounts
        If Incorrect_Account_Length Then
            MsgBox ("Cells with Incorrect Account Number have been hightlighted and sorted for you")
            Incorrect_Account_Length = False
            work_Sheet.Range("B:B").AutoFilter Field:=1, Criteria1:=RGB(204, 255 _
            , 204), Operator:=xlFilterCellColor
        End If
    Next Counter
    'Call Account_Length
End Sub
Private Sub Account_Name()
    Dim Final_Row_Account_Name
    Dim work_Book As Workbook
    Dim work_Sheet As Worksheet
    Dim calculate_Sheet As Worksheet
    Dim final_Sheet As Worksheet
    Set work_Book = ActiveWorkbook
    Set work_Sheet = work_Book.Worksheets("Sheet1")
    Set calculate_Sheet = work_Book.Worksheets("Calculate_Sheet")
    Set final_Sheet = work_Book.Worksheets("Final_Sheet")
    Final_Row_Account_Name = work_Sheet.Cells(Rows.Count, "C").End(xlUp).Row
    
    For Counter = 2 To Final_Row_Account_Name
        calculate_Sheet.Cells(Counter, 1).Value = work_Sheet.Cells(Counter, "C").Value
        calculate_Sheet.Cells(Counter, 2).Value = work_Sheet.Cells(Counter, 1).Value

        If Len(work_Sheet.Cells(Counter, "B").Value) <> 6 And Client_ID <> "55P" Then
            Incorrect_Account_Length = True
            work_Sheet.Cells(Counter, "B").Interior.ColorIndex = 35
        End If
'        Worksheets("Calculate_Sheet").Cells(Counter, 2).Replace What:="`", Replacement:="", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        Worksheets("Calculate_Sheet").Cells(Counter, 2).Replace What:="!", Replacement:="", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        Worksheets("Calculate_Sheet").Cells(Counter, 2).Replace What:="@", Replacement:="AT", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        Worksheets("Calculate_Sheet").Cells(Counter, 2).Replace What:="#", Replacement:="", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        Worksheets("Calculate_Sheet").Cells(Counter, 2).Replace What:="$", Replacement:="", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        Worksheets("Calculate_Sheet").Cells(Counter, 2).Replace What:="%", Replacement:="", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        Worksheets("Calculate_Sheet").Cells(Counter, 2).Replace What:="^", Replacement:="", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        Worksheets("Calculate_Sheet").Cells(Counter, 2).Replace What:="&", Replacement:="AND", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        Worksheets("Calculate_Sheet").Cells(Counter, 2).Replace What:="  ", Replacement:=" ", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        Worksheets("Calculate_Sheet").Cells(Counter, 3).Value = "=IF(OR(RIGHT(TRIM(RC[-1]),1)="")"",RIGHT(TRIM(RC[-1]),1)=""."",RIGHT(TRIM(RC[-1]),1)="",""),LEFT(TRIM(RC[-1]),LEN(TRIM(RC[-1]))-1),TRIM(RC[-1]))"
'
        calculate_Sheet.Cells(Counter, 3).Value = "=LEFT(RC[-2], 50)"
        calculate_Sheet.Cells(Counter, 3).Copy
        calculate_Sheet.Cells(Counter, 4).PasteSpecial xlPasteValues
            
        With calculate_Sheet.Cells(Counter, 4)
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
        'Final Sheet
        calculate_Sheet.Cells(Counter, 4).Copy
        final_Sheet.Cells(Counter, "C").PasteSpecial xlPasteValues
        final_Sheet.Cells(Counter, "C").PasteSpecial xlFormats
    Next Counter
    'Call Account_Length
End Sub
Private Sub Interal_Account_Number()
    Dim Final_Row_Internal_Account_Number
    Dim work_Book As Workbook
    Dim work_Sheet As Worksheet
    Dim calculate_Sheet As Worksheet
    Dim final_Sheet As Worksheet
    Set work_Book = ActiveWorkbook
    Set work_Sheet = work_Book.Worksheets("Sheet1")
    Set calculate_Sheet = work_Book.Worksheets("Calculate_Sheet")
    Set final_Sheet = work_Book.Worksheets("Final_Sheet")
    Final_Row_Internal_Account_Number = work_Sheet.Cells(Rows.Count, "D").End(xlUp).Row
    
    For Counter = 2 To Final_Row_Internal_Account_Number
        calculate_Sheet.Cells(Counter, 1).Value = work_Sheet.Cells(Counter, "D").Value
        calculate_Sheet.Cells(Counter, 2).Value = Worksheets("Calculate_Sheet").Cells(Counter, 1).Value

'        Worksheets("Calculate_Sheet").Cells(Counter, 2).Replace What:="`", Replacement:="", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        Worksheets("Calculate_Sheet").Cells(Counter, 2).Replace What:="!", Replacement:="", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        Worksheets("Calculate_Sheet").Cells(Counter, 2).Replace What:="@", Replacement:="AT", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        Worksheets("Calculate_Sheet").Cells(Counter, 2).Replace What:="#", Replacement:="", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        Worksheets("Calculate_Sheet").Cells(Counter, 2).Replace What:="$", Replacement:="", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        Worksheets("Calculate_Sheet").Cells(Counter, 2).Replace What:="%", Replacement:="", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        Worksheets("Calculate_Sheet").Cells(Counter, 2).Replace What:="^", Replacement:="", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        Worksheets("Calculate_Sheet").Cells(Counter, 2).Replace What:="&", Replacement:="AND", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
        calculate_Sheet.Cells(Counter, 2).Replace What:="  ", Replacement:=" ", LookAt:=xlPart, SearchOrder:= _
            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

        calculate_Sheet.Cells(Counter, 3).Value = "=IF(OR(RIGHT(TRIM(RC[-1]),1)="")"",RIGHT(TRIM(RC[-1]),1)=""."",RIGHT(TRIM(RC[-1]),1)="",""),LEFT(TRIM(RC[-1]),LEN(TRIM(RC[-1]))-1),TRIM(RC[-1]))"

        calculate_Sheet.Cells(Counter, 3).Copy
        calculate_Sheet.Cells(Counter, 4).PasteSpecial xlPasteValues
            
        With calculate_Sheet.Cells(Counter, 4)
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
        'Final Sheet
        calculate_Sheet.Cells(Counter, 4).Copy
        final_Sheet.Cells(Counter, "D").PasteSpecial xlPasteValues
        final_Sheet.Cells(Counter, "D").PasteSpecial xlFormats
        'Filter Incorrect Accounts
'        If Incorrect_Account_Length Then
'            MsgBox ("Cells with Incorrect Account Length have been hightlighted and sorted for you")
'            Incorrect_Account_Length = False
'            work_Sheet.Range("D:D").AutoFilter Field:=1, Criteria1:=RGB(204, 255 _
'            , 204), Operator:=xlFilterCellColor
'        End If
    Next Counter
    'Call Account_Length
End Sub
Private Sub Check_PEID()
    Dim Final_Row_PEID
    Dim work_Book As Workbook
    Dim work_Sheet As Worksheet
    Dim calculate_Sheet As Worksheet
    Dim final_Sheet As Worksheet
    Set work_Book = ActiveWorkbook
    Set work_Sheet = work_Book.Worksheets("Sheet1")
    Set calculate_Sheet = work_Book.Worksheets("Calculate_Sheet")
    Set final_Sheet = work_Book.Worksheets("Final_Sheet")
    Dim Incorrect_Account_Length As Boolean
    Final_Row_PEID = work_Sheet.Cells(Rows.Count, "E").End(xlUp).Row
    
    For Counter = 2 To Final_Row_PEID
        calculate_Sheet.Cells(Counter, 1).Value = work_Sheet.Cells(Counter, "E").Value
        calculate_Sheet.Cells(Counter, 2).Value = Worksheets("Calculate_Sheet").Cells(Counter, 1).Value

        If Len(work_Sheet.Cells(Counter, "E").Value) <> 3 Then
            Incorrect_Account_Length = True
            work_Sheet.Cells(Counter, "E").Interior.ColorIndex = 35
        End If
'        Worksheets("Calculate_Sheet").Cells(Counter, 2).Replace What:="`", Replacement:="", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        Worksheets("Calculate_Sheet").Cells(Counter, 2).Replace What:="!", Replacement:="", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        Worksheets("Calculate_Sheet").Cells(Counter, 2).Replace What:="@", Replacement:="AT", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        Worksheets("Calculate_Sheet").Cells(Counter, 2).Replace What:="#", Replacement:="", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        Worksheets("Calculate_Sheet").Cells(Counter, 2).Replace What:="$", Replacement:="", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        Worksheets("Calculate_Sheet").Cells(Counter, 2).Replace What:="%", Replacement:="", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        Worksheets("Calculate_Sheet").Cells(Counter, 2).Replace What:="^", Replacement:="", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        Worksheets("Calculate_Sheet").Cells(Counter, 2).Replace What:="&", Replacement:="AND", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
        calculate_Sheet.Cells(Counter, 2).Replace What:="  ", Replacement:=" ", LookAt:=xlPart, SearchOrder:= _
            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

        calculate_Sheet.Cells(Counter, 3).Value = "=IF(OR(RIGHT(TRIM(RC[-1]),1)="")"",RIGHT(TRIM(RC[-1]),1)=""."",RIGHT(TRIM(RC[-1]),1)="",""),LEFT(TRIM(RC[-1]),LEN(TRIM(RC[-1]))-1),TRIM(RC[-1]))"

        calculate_Sheet.Cells(Counter, 3).Copy
        calculate_Sheet.Cells(Counter, 4).PasteSpecial xlPasteValues
            
        With calculate_Sheet.Cells(Counter, 4)
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
        'Final Sheet
        calculate_Sheet.Cells(Counter, 4).Copy
        final_Sheet.Cells(Counter, "E").PasteSpecial xlPasteValues
        final_Sheet.Cells(Counter, "E").PasteSpecial xlFormats
        'Filter Incorrect Accounts
        If Incorrect_Account_Length Then
            MsgBox ("Cells with Incorrect PEID have been hightlighted and sorted for you")
            Incorrect_Account_Length = False
            work_Sheet.Range("E:E").AutoFilter Field:=1, Criteria1:=RGB(204, 255 _
            , 204), Operator:=xlFilterCellColor
        End If
    Next Counter
    'Call Account_Length
End Sub
Private Sub Account_Flags()
    Dim Final_Row_Account_Flags
    Dim work_Book As Workbook
    Dim work_Sheet As Worksheet
    Dim final_Sheet As Worksheet
    Set work_Book = ActiveWorkbook
    Set work_Sheet = work_Book.Worksheets("Sheet1")
    Set final_Sheet = work_Book.Worksheets("Final_Sheet")
    Final_Row_Account_Flags = work_Sheet.Cells(Rows.Count, 1).End(xlUp).Row
    
    final_Sheet.Range(Cells(Final_Row_Account_Flags, "P"), Cells(2, "P")).Value = "Y"
    final_Sheet.Range(Cells(Final_Row_Account_Flags, "Q"), Cells(2, "Q")).Value = "S"
    final_Sheet.Range(Cells(Final_Row_Account_Flags, "S"), Cells(2, "S")).Value = "B"
End Sub
