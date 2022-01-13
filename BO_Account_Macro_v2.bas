Attribute VB_Name = "Module2"
    'BO_Adder by Diego Espitia, January 2022
    Global work_Sheet As Worksheet
    Global work_Book As ThisWorkbook
    Global incorrect_Account_Length As Boolean
    Public final_Row_Data
    Option Compare Text
Sub Main() 'Main BO_Adder
Attribute Main.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim cell_Data As Range
    Dim data_Range As Range
    incorrect_Acount_Length = False

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    'Call check_Columns_For_Text //DON'T ENABLE//
    Create_Final_Sheet
    Check_CID
    Account_Number
    
    If incorrect_Account_Length Then
        Exit Sub
    End If
    
    Entry_Type
    Market_ID
    Entity_Type_Code
    BO_Name
    URN_Type
    URN_Number
    Address_Type
    Address
    'Address_2
    'Address_3
    'Address_4
    'Address_5
    Street_Name
    Building_No
    PO_Box
    Postal_Code
    Town
    Province
    Country_Code
    
    Set work_Sheet = Worksheets("Sheet1")
    Set final_Sheet = Worksheets("Final_Sheet")
    
    work_Sheet.Rows("1:1").Copy
    final_Sheet.Rows("1:1").PasteSpecial xlPasteColumnWidths
    final_Sheet.Rows("1:1").PasteSpecial xlPasteValues
    final_Sheet.Rows("1:1").PasteSpecial xlFormats
    With final_Sheet.Range(final_Sheet.Cells(2, "A"), final_Sheet.Cells(final_Row_Data, "AA")).Borders
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
    'Dim final_Row_Data
    Dim range_Client_ID
    Dim work_Book As Workbook
    Dim work_Sheet As Worksheet
    'Dim calculate_Sheet As Worksheet
    Dim final_Sheet As Worksheet
    Set work_Book = ThisWorkbook
    Set work_Sheet = work_Book.Worksheets("Sheet1")
    'Set calculate_Sheet = work_Book.Worksheets("Calculate_Sheet")
    Set final_Sheet = work_Book.Worksheets("Final_Sheet")
    final_Row_Data = work_Sheet.Cells(Rows.Count, "B").End(xlUp).Row
    
    'final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "B"), final_Sheet.Cells(2, "B")).Value = _
    '    work_Sheet.Range(work_Sheet.Cells(final_Row_Data, "B"), work_Sheet.Cells(3, "B")).Value
    
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
        'final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "B"), final_Sheet.Cells(2, "B")).PasteSpecial xlPasteFormats
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "C"), final_Sheet.Cells(2, "C")).ClearContents
        
'        With final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "B"), final_Sheet.Cells(2, "B"))
'            .HorizontalAlignment = xlLeft
'            .VerticalAlignment = xlBottom
'            .WrapText = False
'            .Orientation = 0
'            .AddIndent = False
'            .IndentLevel = 0
'            .ShrinkToFit = False
'            .ReadingOrder = xlContext
'            .MergeCells = False
'        End With
        
        Account_Length
    End If
End Sub
Private Sub Entry_Type()
    'Dim Final_Row_Entity_Code_ID
    Dim work_Book As Workbook
    Dim work_Sheet As Worksheet
    'Dim calculate_Sheet As Worksheet
    Dim final_Sheet As Worksheet
    Set work_Book = ThisWorkbook
    Set work_Sheet = work_Book.Worksheets("Sheet1")
    'Set calculate_Sheet = work_Book.Worksheets("Calculate_Sheet")
    Set final_Sheet = work_Book.Worksheets("Final_Sheet")
    final_Row_Data = work_Sheet.Cells(Rows.Count, 1).End(xlUp).Row
    
    work_Sheet.Range(work_Sheet.Cells(final_Row_Data, "C"), work_Sheet.Cells(2, "C")).Copy
    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "C"), final_Sheet.Cells(2, "C")).PasteSpecial xlPasteValues
    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "C"), final_Sheet.Cells(2, "C")).PasteSpecial xlPasteFormats
    
    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "C"), final_Sheet.Cells(2, "C")).Replace What:="`,!,@,#,$,%,^", Replacement:="", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        
    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "C"), final_Sheet.Cells(2, "C")).Replace What:="@", Replacement:="AT", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        
    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "C"), final_Sheet.Cells(2, "C")).Replace What:="&", Replacement:="AND", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        
    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "C"), final_Sheet.Cells(2, "C")).Replace What:="  ", Replacement:=" ", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        
    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "D"), final_Sheet.Cells(2, "D")).Value = _
        "=IF(OR(RIGHT(TRIM(RC[-1]),1)="")"",RIGHT(TRIM(RC[-1]),1)=""."",RIGHT(TRIM(RC[-1]),1)="",""),LEFT(TRIM(RC[-1]),LEN(TRIM(RC[-1]))-1),TRIM(RC[-1]))"
    
    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "D"), final_Sheet.Cells(2, "D")).Copy
    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "C"), final_Sheet.Cells(2, "C")).PasteSpecial xlPasteValues
    'final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "C"), final_Sheet.Cells(2, "C")).PasteSpecial xlPasteFormats
    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "D"), final_Sheet.Cells(2, "D")).ClearContents
    
'    With final_Sheet.Cells(final_Row_Data, "C")
'        .HorizontalAlignment = xlLeft
'        .VerticalAlignment = xlBottom
'        .WrapText = False
'        .Orientation = 0
'        .AddIndent = False
'        .IndentLevel = 0
'        .ShrinkToFit = False
'        .ReadingOrder = xlContext
'        .MergeCells = False
'    End With
End Sub
Private Sub Market_ID() 'Fill in Markets Column
    Dim Final_Row_Market_ID
    Dim work_Book As Workbook
    Dim work_Sheet As Worksheet
    'Dim calculate_Sheet As Worksheet
    Dim final_Sheet As Worksheet
    Set work_Book = ThisWorkbook
    Set work_Sheet = work_Book.Worksheets("Sheet1")
    'Set calculate_Sheet = work_Book.Worksheets("Calculate_Sheet")
    Set final_Sheet = work_Book.Worksheets("Final_Sheet")
    Final_Row_Market_ID = work_Sheet.Cells(Rows.Count, 1).End(xlUp).Row
    
    final_Sheet.Select
    final_Sheet.Range(final_Sheet.Cells(Final_Row_Market_ID, "D"), final_Sheet.Cells(2, "D")).Value = "***"
    work_Sheet.Select
End Sub
Private Sub Entity_Type_Code()
    'Dim Final_Row_Entity_Code_ID
    Dim work_Book As Workbook
    Dim work_Sheet As Worksheet
    'Dim calculate_Sheet As Worksheet
    Dim final_Sheet As Worksheet
    Set work_Book = ThisWorkbook
    Set work_Sheet = work_Book.Worksheets("Sheet1")
    'Set calculate_Sheet = work_Book.Worksheets("Calculate_Sheet")
    Set final_Sheet = work_Book.Worksheets("Final_Sheet")
    final_Row_Data = work_Sheet.Cells(Rows.Count, 1).End(xlUp).Row
    
    work_Sheet.Range(work_Sheet.Cells(final_Row_Data, "E"), work_Sheet.Cells(2, "E")).Copy
    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "E"), final_Sheet.Cells(2, "E")).PasteSpecial xlPasteValues
    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "E"), final_Sheet.Cells(2, "E")).PasteSpecial xlPasteFormats
    
'    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "E"), final_Sheet.Cells(2, "E")).Replace What:="`,!,@,#,$,%,^", Replacement:="", LookAt:= _
'        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "E"), final_Sheet.Cells(2, "E")).Replace What:="@", Replacement:="AT", LookAt:= _
'        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "E"), final_Sheet.Cells(2, "E")).Replace What:="&", Replacement:="AND", LookAt:= _
'        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "E"), final_Sheet.Cells(2, "E")).Replace What:="  ", Replacement:=" ", LookAt:= _
'        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        
    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "F"), final_Sheet.Cells(2, "F")).Value = _
        "=IF(OR(RIGHT(TRIM(RC[-1]),1)="")"",RIGHT(TRIM(RC[-1]),1)=""."",RIGHT(TRIM(RC[-1]),1)="",""),LEFT(TRIM(RC[-1]),LEN(TRIM(RC[-1]))-1),TRIM(RC[-1]))"
    
    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "F"), final_Sheet.Cells(2, "F")).Copy
    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "E"), final_Sheet.Cells(2, "E")).PasteSpecial xlPasteValues
    'final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "E"), final_Sheet.Cells(2, "E")).PasteSpecial xlPasteFormats
    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "F"), final_Sheet.Cells(2, "F")).ClearContents
    
'    For Counter = 3 To final_Row_Data
'        calculate_Sheet.Cells(Counter, 1).Value = work_Sheet.Cells(Counter, "E").Value
'        calculate_Sheet.Cells(Counter, 2).Value = calculate_Sheet.Cells(Counter, 1).Value
'
'        calculate_Sheet.Cells(Counter, 2).Replace What:="`", Replacement:="", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        calculate_Sheet.Cells(Counter, 2).Replace What:="!", Replacement:="", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        calculate_Sheet.Cells(Counter, 2).Replace What:="@", Replacement:="AT", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        calculate_Sheet.Cells(Counter, 2).Replace What:="#", Replacement:="", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        calculate_Sheet.Cells(Counter, 2).Replace What:="$", Replacement:="", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        calculate_Sheet.Cells(Counter, 2).Replace What:="%", Replacement:="", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        calculate_Sheet.Cells(Counter, 2).Replace What:="^", Replacement:="", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        calculate_Sheet.Cells(Counter, 2).Replace What:="&", Replacement:="AND", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        calculate_Sheet.Cells(Counter, 2).Replace What:="  ", Replacement:=" ", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        calculate_Sheet.Cells(Counter, 3).Value = "=IF(OR(RIGHT(TRIM(RC[-1]),1)="")"",RIGHT(TRIM(RC[-1]),1)=""."",RIGHT(TRIM(RC[-1]),1)="",""),LEFT(TRIM(RC[-1]),LEN(TRIM(RC[-1]))-1),TRIM(RC[-1]))"
'
'        calculate_Sheet.Cells(Counter, 3).Copy
'        final_Sheet.Cells(Counter - 1, "E").PasteSpecial xlPasteValues
'
'        With final_Sheet.Cells(final_Row_Data, "E")
'            .HorizontalAlignment = xlLeft
'            .VerticalAlignment = xlBottom
'            .WrapText = False
'            .Orientation = 0
'            .AddIndent = False
'            .IndentLevel = 0
'            .ShrinkToFit = False
'            .ReadingOrder = xlContext
'            .MergeCells = False
'        End With
'        'Final Sheet
'        'Worksheets("Final_Sheet").Cells(Counter - 1, "E").Value = calculate_Sheet.Cells(Counter, 4).Value
'    Next Counter
End Sub
Private Sub Check_CID() 'Checking for CID
    'Dim final_Row_Data
    Dim work_Book As Workbook
    Dim work_Sheet As Worksheet
    'Dim calculate_Sheet As Worksheet
    Dim final_Sheet As Worksheet
    Set work_Book = ThisWorkbook
    Set work_Sheet = work_Book.Worksheets("Sheet1")
    'Set calculate_Sheet = work_Book.Worksheets("Calculate_Sheet")
    Set final_Sheet = work_Book.Worksheets("Final_Sheet")
    final_Row_Data = work_Sheet.Cells(Rows.Count, 1).End(xlUp).Row
    
    work_Sheet.Range(work_Sheet.Cells(final_Row_Data, "A"), work_Sheet.Cells(3, "A")).Copy
    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "A"), final_Sheet.Cells(2, "A")).PasteSpecial xlPasteValues
    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "A"), final_Sheet.Cells(2, "A")).PasteSpecial xlPasteFormats
    
'    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "A"), final_Sheet.Cells(2, "A")).Replace What:="`,!,@,#,$,%,^", Replacement:="", LookAt:= _
'        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "A"), final_Sheet.Cells(2, "A")).Replace What:="@", Replacement:="AT", LookAt:= _
'        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "A"), final_Sheet.Cells(2, "A")).Replace What:="&", Replacement:="AND", LookAt:= _
'        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "A"), final_Sheet.Cells(2, "A")).Replace What:="  ", Replacement:=" ", LookAt:= _
'        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        
    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "B"), final_Sheet.Cells(2, "B")).Value = _
        "=IF(OR(RIGHT(TRIM(RC[-1]),1)="")"",RIGHT(TRIM(RC[-1]),1)=""."",RIGHT(TRIM(RC[-1]),1)="",""),LEFT(TRIM(RC[-1]),LEN(TRIM(RC[-1]))-1),TRIM(RC[-1]))"
    
    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "B"), final_Sheet.Cells(2, "B")).Copy
    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "A"), final_Sheet.Cells(2, "A")).PasteSpecial xlPasteValues
    'final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "A"), final_Sheet.Cells(2, "A")).PasteSpecial xlPasteFormats
    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "B"), final_Sheet.Cells(2, "B")).ClearContents
    
'    For Counter = 3 To final_Row_Data
'        calculate_Sheet.Cells(Counter, 1).Value = work_Sheet.Cells(Counter, "A").Value
'        calculate_Sheet.Cells(Counter, 2).Value = calculate_Sheet.Cells(Counter, 1).Value
'
'        calculate_Sheet.Cells(Counter, 2).Replace What:="`", Replacement:="", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        calculate_Sheet.Cells(Counter, 2).Replace What:="!", Replacement:="", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        calculate_Sheet.Cells(Counter, 2).Replace What:="@", Replacement:="AT", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        calculate_Sheet.Cells(Counter, 2).Replace What:="#", Replacement:="", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        calculate_Sheet.Cells(Counter, 2).Replace What:="$", Replacement:="", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        calculate_Sheet.Cells(Counter, 2).Replace What:="%", Replacement:="", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        calculate_Sheet.Cells(Counter, 2).Replace What:="^", Replacement:="", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        calculate_Sheet.Cells(Counter, 2).Replace What:="&", Replacement:="AND", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        calculate_Sheet.Cells(Counter, 2).Replace What:="  ", Replacement:=" ", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        calculate_Sheet.Cells(Counter, 3).Value = "=IF(OR(RIGHT(TRIM(RC[-1]),1)="")"",RIGHT(TRIM(RC[-1]),1)=""."",RIGHT(TRIM(RC[-1]),1)="",""),LEFT(TRIM(RC[-1]),LEN(TRIM(RC[-1]))-1),TRIM(RC[-1]))"
'
'        calculate_Sheet.Cells(Counter, 3).Copy
'        final_Sheet.Cells(Counter - 1, "A").PasteSpecial xlPasteValues
'
'        With final_Sheet.Cells(final_Row_Data, "A")
'            .HorizontalAlignment = xlLeft
'            .VerticalAlignment = xlBottom
'            .WrapText = False
'            .Orientation = 0
'            .AddIndent = False
'            .IndentLevel = 0
'            .ShrinkToFit = False
'            .ReadingOrder = xlContext
'            .MergeCells = False
'        End With
'        'Final Sheet
'        'final_Sheet.Cells(Counter - 1, "A").Value = calculate_Sheet.Cells(Counter, 4).Value
'    Next Counter
End Sub
Private Sub BO_Name() 'Check for BO Name
    'Dim final_Row_Data
    Dim work_Book As Workbook
    Dim work_Sheet As Worksheet
    'Dim calculate_Sheet As Worksheet
    Dim final_Sheet As Worksheet
    Set work_Book = ThisWorkbook
    Set work_Sheet = work_Book.Worksheets("Sheet1")
    'Set calculate_Sheet = work_Book.Worksheets("Calculate_Sheet")
    Set final_Sheet = work_Book.Worksheets("Final_Sheet")
    final_Row_Data = work_Sheet.Cells(Rows.Count, "H").End(xlUp).Row
    
    work_Sheet.Range(work_Sheet.Cells(final_Row_Data, "H"), work_Sheet.Cells(2, "H")).Copy
    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "H"), final_Sheet.Cells(2, "H")).PasteSpecial xlPasteValues
    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "H"), final_Sheet.Cells(2, "H")).PasteSpecial xlPasteFormats
    
'    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "H"), final_Sheet.Cells(2, "H")).Replace What:="`,!,@,#,$,%,^", Replacement:="", LookAt:= _
'        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "H"), final_Sheet.Cells(2, "H")).Replace What:="@", Replacement:="AT", LookAt:= _
'        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "H"), final_Sheet.Cells(2, "H")).Replace What:="&", Replacement:="AND", LookAt:= _
'        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "H"), final_Sheet.Cells(2, "H")).Replace What:="  ", Replacement:=" ", LookAt:= _
'        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        
    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "I"), final_Sheet.Cells(2, "I")).Value = _
        "=IF(OR(RIGHT(TRIM(RC[-1]),1)="")"",RIGHT(TRIM(RC[-1]),1)=""."",RIGHT(TRIM(RC[-1]),1)="",""),LEFT(TRIM(RC[-1]),LEN(TRIM(RC[-1]))-1),TRIM(RC[-1]))"
    
    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "I"), final_Sheet.Cells(2, "I")).Copy
    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "H"), final_Sheet.Cells(2, "H")).PasteSpecial xlPasteValues
    'final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "H"), final_Sheet.Cells(2, "H")).PasteSpecial xlPasteFormats
    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "I"), final_Sheet.Cells(2, "I")).ClearContents
    
'    For Counter = 3 To Final_Row
'
'        calculate_Sheet.Cells(Counter, 1).Value = work_Sheet.Cells(Counter, "H").Value
'        calculate_Sheet.Cells(Counter, 2).Value = calculate_Sheet.Cells(Counter, 1).Value
'
'        calculate_Sheet.Cells(Counter, 2).Replace What:="`", Replacement:="", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        calculate_Sheet.Cells(Counter, 2).Replace What:="!", Replacement:="", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        calculate_Sheet.Cells(Counter, 2).Replace What:="@", Replacement:="AT", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        calculate_Sheet.Cells(Counter, 2).Replace What:="#", Replacement:="", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        calculate_Sheet.Cells(Counter, 2).Replace What:="$", Replacement:="", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        calculate_Sheet.Cells(Counter, 2).Replace What:="%", Replacement:="", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        calculate_Sheet.Cells(Counter, 2).Replace What:="^", Replacement:="", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        calculate_Sheet.Cells(Counter, 2).Replace What:="&", Replacement:="AND", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        calculate_Sheet.Cells(Counter, 2).Replace What:="  ", Replacement:=" ", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        calculate_Sheet.Cells(Counter, 3).Value = "=IF(OR(RIGHT(TRIM(RC[-1]),1)="")"",RIGHT(TRIM(RC[-1]),1)=""."",RIGHT(TRIM(RC[-1]),1)="",""),LEFT(TRIM(RC[-1]),LEN(TRIM(RC[-1]))-1),TRIM(RC[-1]))"
'
'        calculate_Sheet.Cells(Counter, 3).Copy
'        final_Sheet.Cells(Counter - 1, "H").PasteSpecial xlPasteValues
            
'        With final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "H"), final_Sheet.Cells(2, "H"))
'            .HorizontalAlignment = xlLeft
'            .VerticalAlignment = xlBottom
'            .WrapText = False
'            .Orientation = 0
'            .AddIndent = False
'            .IndentLevel = 0
'            .ShrinkToFit = False
'            .ReadingOrder = xlContext
'            .MergeCells = False
'        End With
'        'final_Sheet.Cells(Counter - 1, "H").Value = calculate_Sheet.Cells(Counter, 4).Value
'    Next Counter
End Sub
Private Sub URN_Type()
    'Dim Final_Row_URN_Type
    Dim work_Book As Workbook
    Dim work_Sheet As Worksheet
    'Dim calculate_Sheet As Worksheet
    Dim final_Sheet As Worksheet
    Set work_Book = ThisWorkbook
    Set work_Sheet = work_Book.Worksheets("Sheet1")
    'Set calculate_Sheet = work_Book.Worksheets("Calculate_Sheet")
    Set final_Sheet = work_Book.Worksheets("Final_Sheet")
    final_Row_Data = work_Sheet.Cells(Rows.Count, "J").End(xlUp).Row
    
    If final_Row_Data <> 1 Then 'Might have to loop through URN columns to find if N/A and clear it
        work_Sheet.Range(work_Sheet.Cells(final_Row_Data, "J"), work_Sheet.Cells(2, "J")).Copy
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "J"), final_Sheet.Cells(2, "J")).PasteSpecial xlPasteValues
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "J"), final_Sheet.Cells(2, "J")).PasteSpecial xlPasteFormats
        
'        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "J"), final_Sheet.Cells(2, "J")).Replace What:="`,!,@,#,$,%,^", Replacement:="", LookAt:= _
'            xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "J"), final_Sheet.Cells(2, "J")).Replace What:="@", Replacement:="AT", LookAt:= _
'            xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "J"), final_Sheet.Cells(2, "J")).Replace What:="&", Replacement:="AND", LookAt:= _
'            xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "J"), final_Sheet.Cells(2, "J")).Replace What:="  ", Replacement:=" ", LookAt:= _
'            xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "J"), final_Sheet.Cells(2, "J")).Replace What:="N/A", Replacement:="", LookAt:= _
            xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
            
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "K"), final_Sheet.Cells(2, "K")).Value = _
            "=IF(OR(RIGHT(TRIM(RC[-1]),1)="")"",RIGHT(TRIM(RC[-1]),1)=""."",RIGHT(TRIM(RC[-1]),1)="",""),LEFT(TRIM(RC[-1]),LEN(TRIM(RC[-1]))-1),TRIM(RC[-1]))"
        
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "K"), final_Sheet.Cells(2, "K")).Copy
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "J"), final_Sheet.Cells(2, "J")).PasteSpecial xlPasteValues
        'final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "J"), final_Sheet.Cells(2, "J")).PasteSpecial xlPasteFormats
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "K"), final_Sheet.Cells(2, "K")).ClearContents
        
'        With final_Sheet.Cells(final_Row_Data, "J")
'            .HorizontalAlignment = xlLeft
'            .VerticalAlignment = xlBottom
'            .WrapText = False
'            .Orientation = 0
'            .AddIndent = False
'            .IndentLevel = 0
'            .ShrinkToFit = False
'            .ReadingOrder = xlContext
'            .MergeCells = False
'        End With
    End If
End Sub
Private Sub URN_Number()
    'Dim final_Row_URN_Number
    Dim work_Book As Workbook
    Dim work_Sheet As Worksheet
    'Dim calculate_Sheet As Worksheet
    Dim final_Sheet As Worksheet
    Set work_Book = ThisWorkbook
    Set work_Sheet = work_Book.Worksheets("Sheet1")
    'Set calculate_Sheet = work_Book.Worksheets("Calculate_Sheet")
    Set final_Sheet = work_Book.Worksheets("Final_Sheet")
    final_Row_Data = work_Sheet.Cells(Rows.Count, "K").End(xlUp).Row
    
    If final_Row_Data <> 1 Then
        work_Sheet.Range(work_Sheet.Cells(final_Row_Data, "K"), work_Sheet.Cells(2, "K")).Copy
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "K"), final_Sheet.Cells(2, "K")).PasteSpecial xlPasteValues
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "K"), final_Sheet.Cells(2, "K")).PasteSpecial xlPasteFormats
        
'        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "K"), final_Sheet.Cells(2, "K")).Replace What:="`,!,@,#,$,%,^", Replacement:="", LookAt:= _
'            xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "K"), final_Sheet.Cells(2, "K")).Replace What:="@", Replacement:="AT", LookAt:= _
'            xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "K"), final_Sheet.Cells(2, "K")).Replace What:="&", Replacement:="AND", LookAt:= _
'            xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "K"), final_Sheet.Cells(2, "K")).Replace What:="  ", Replacement:=" ", LookAt:= _
'            xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "K"), final_Sheet.Cells(2, "K")).Replace What:="N/A", Replacement:="", LookAt:= _
            xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
            
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "L"), final_Sheet.Cells(2, "L")).Value = _
            "=IF(OR(RIGHT(TRIM(RC[-1]),1)="")"",RIGHT(TRIM(RC[-1]),1)=""."",RIGHT(TRIM(RC[-1]),1)="",""),LEFT(TRIM(RC[-1]),LEN(TRIM(RC[-1]))-1),TRIM(RC[-1]))"
        
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "L"), final_Sheet.Cells(2, "L")).Copy
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "K"), final_Sheet.Cells(2, "K")).PasteSpecial xlPasteValues
        'final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "K"), final_Sheet.Cells(2, "K")).PasteSpecial xlPasteFormats
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "L"), final_Sheet.Cells(2, "L")).ClearContents
        
'        With final_Sheet.Cells(final_Row_Data, "K")
'            .HorizontalAlignment = xlLeft
'            .VerticalAlignment = xlBottom
'            .WrapText = False
'            .Orientation = 0
'            .AddIndent = False
'            .IndentLevel = 0
'            .ShrinkToFit = False
'            .ReadingOrder = xlContext
'            .MergeCells = False
'        End With
    End If
End Sub
Private Sub Address_Type()
    'Dim Final_Row_Address_Type
    Dim work_Book As Workbook
    Dim work_Sheet As Worksheet
    'Dim calculate_Sheet As Worksheet
    Dim final_Sheet As Worksheet
    Set work_Book = ThisWorkbook
    Set work_Sheet = work_Book.Worksheets("Sheet1")
    'Set calculate_Sheet = work_Book.Worksheets("Calculate_Sheet")
    Set final_Sheet = work_Book.Worksheets("Final_Sheet")
    final_Row_Data = work_Sheet.Cells(Rows.Count, "O").End(xlUp).Row
    
    If final_Row_Data <> 1 Then
        work_Sheet.Range(work_Sheet.Cells(final_Row_Data, "O"), work_Sheet.Cells(2, "O")).Copy
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "O"), final_Sheet.Cells(2, "O")).PasteSpecial xlPasteValues
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "O"), final_Sheet.Cells(2, "O")).PasteSpecial xlPasteFormats
        
'        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "O"), final_Sheet.Cells(2, "O")).Replace What:="`,!,@,#,$,%,^", Replacement:="", LookAt:= _
'            xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "O"), final_Sheet.Cells(2, "O")).Replace What:="@", Replacement:="AT", LookAt:= _
'            xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "O"), final_Sheet.Cells(2, "O")).Replace What:="&", Replacement:="AND", LookAt:= _
'            xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "O"), final_Sheet.Cells(2, "O")).Replace What:="  ", Replacement:=" ", LookAt:= _
'            xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
            
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "P"), final_Sheet.Cells(2, "P")).Value = _
            "=IF(OR(RIGHT(TRIM(RC[-1]),1)="")"",RIGHT(TRIM(RC[-1]),1)=""."",RIGHT(TRIM(RC[-1]),1)="",""),LEFT(TRIM(RC[-1]),LEN(TRIM(RC[-1]))-1),TRIM(RC[-1]))"
        
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "P"), final_Sheet.Cells(2, "P")).Copy
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "O"), final_Sheet.Cells(2, "O")).PasteSpecial xlPasteValues
        'final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "O"), final_Sheet.Cells(2, "O")).PasteSpecial xlPasteFormats
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "P"), final_Sheet.Cells(2, "P")).ClearContents
        
'        With final_Sheet.Cells(final_Row_Data, "O")
'            .HorizontalAlignment = xlLeft
'            .VerticalAlignment = xlBottom
'            .WrapText = False
'            .Orientation = 0
'            .AddIndent = False
'            .IndentLevel = 0
'            .ShrinkToFit = False
'            .ReadingOrder = xlContext
'            .MergeCells = False
'        End With
    End If
End Sub
Private Sub Address()
    'Dim Final_Row_Address
    Dim work_Book As Workbook
    Dim work_Sheet As Worksheet
    'Dim calculate_Sheet As Worksheet
    Dim final_Sheet As Worksheet
    Set work_Book = ThisWorkbook
    Set work_Sheet = work_Book.Worksheets("Sheet1")
    'Set calculate_Sheet = work_Book.Worksheets("Calculate_Sheet")
    Set final_Sheet = work_Book.Worksheets("Final_Sheet")
    Dim account_Address_1_Range As Range
    Dim account_Address_2_Range As Range
    Dim account_Address_3_Range As Range
    Dim account_Address_4_Range As Range
    Dim account_Address_5_Range As Range
    'illegal_Characters = Array("`", "!", "#", "$", "%", "^")
    Dim column_A_To_AA As Range
    Dim account_Address_Cell_1 As Range
    Dim account_Address_Cell_2 As Range
    Dim account_Address_Cell_3 As Range
    Dim account_Address_Cell_4 As Range
    Dim account_Address_Cell_5 As Range
    final_Row_Data = work_Sheet.Cells(Rows.Count, "P").End(xlUp).Row 'final_Row for Address 1
    Set account_Address_1_Range = work_Sheet.Range(work_Sheet.Cells(2, "P"), work_Sheet.Cells(final_Row_Data, "P")) 'Set up Address Range 1 on Sheet1
    
'    Dim s As Range
'
'    Set s = Worksheets("Final_Sheet").Cells(2, "P")
'    junk = Array("q", "w", "e", "r", "t", "y", "!")
'
'    For Each a In junk
'        s = Replace(s, a, "")
'        MsgBox (s)
'    Next a
    
    If final_Row_Data <> 1 Then 'Address 1 final_Row_Data cannot be 1
        work_Sheet.Range(work_Sheet.Cells(final_Row_Data, "P"), work_Sheet.Cells(2, "P")).Copy
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "P"), final_Sheet.Cells(2, "P")).PasteSpecial xlPasteValues
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "P"), final_Sheet.Cells(2, "P")).PasteSpecial xlPasteFormats
        
'        Set account_Address_1_Range = final_Sheet.Range(final_Sheet.Cells(2, "P"), final_Sheet.Cells(final_Row_Data, "P")) 'Set up Address Range 1 on Final_Sheet
'
'        For Each account_Address_Cell_1 In account_Address_1_Range.Cells
'            For Each invalid_Data In illegal_Characters
'                account_Address_Cell_1 = Replace(account_Address_Cell_1, invalid_Data, "")
'                account_Address_Cell_1 = Replace(account_Address_Cell_1, "@", "AT")
'                account_Address_Cell_1 = Replace(account_Address_Cell_1, "  ", " ")
'                account_Address_Cell_1 = Replace(account_Address_Cell_1, "&", "AND")
'            Next invalid_Data
'        Next account_Address_Cell_1
'        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "P"), final_Sheet.Cells(2, "P")).Replace What:="`,!,@,#,$,%,^", Replacement:="", LookAt:= _
'            xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "P"), final_Sheet.Cells(2, "P")).Replace What:="!", Replacement:="", LookAt:= _
'            xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "P"), final_Sheet.Cells(2, "P")).Replace What:="@", Replacement:="AT", LookAt:= _
'            xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "P"), final_Sheet.Cells(2, "P")).Replace What:="$", Replacement:="", LookAt:= _
'            xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "P"), final_Sheet.Cells(2, "P")).Replace What:="&", Replacement:="AND", LookAt:= _
'            xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "P"), final_Sheet.Cells(2, "P")).Replace What:="  ", Replacement:=" ", LookAt:= _
'            xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
            
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "Q"), final_Sheet.Cells(2, "Q")).Value = _
            "=IF(OR(RIGHT(TRIM(RC[-1]),1)="")"",RIGHT(TRIM(RC[-1]),1)=""."",RIGHT(TRIM(RC[-1]),1)="",""),LEFT(TRIM(RC[-1]),LEN(TRIM(RC[-1]))-1),TRIM(RC[-1]))"
        
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "Q"), final_Sheet.Cells(2, "Q")).Copy
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "P"), final_Sheet.Cells(2, "P")).PasteSpecial xlPasteValues
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "P"), final_Sheet.Cells(2, "P")).PasteSpecial xlPasteFormats
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "Q"), final_Sheet.Cells(2, "Q")).ClearContents
        
        For Each account_Address_Cell_1 In account_Address_1_Range.Cells 'Address 1 Text to Columns
            If Len(account_Address_Cell_1.Value) > 70 Then
                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_1.Row, "P"), final_Sheet.Cells(account_Address_Cell_1.Row, "P")).TextToColumns Destination:= _
                    final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_1.Row, "P"), final_Sheet.Cells(account_Address_Cell_1.Row, "P")), DataType:=xlFixedWidth, _
                        FieldInfo:=Array(Array(0, 1), Array(70, 1)), TrailingMinusNumbers:=True
            End If
        Next account_Address_Cell_1
        
    final_Row_Data = work_Sheet.Cells(Rows.Count, "Q").End(xlUp).Row 'final_Row for Address 2
    
    If final_Row_Data <> 1 Then 'Address 2, final_Row_Data cannot be 1
        Set account_Address_2_Range = final_Sheet.Range(final_Sheet.Cells(2, "Q"), final_Sheet.Cells(final_Row_Data, "Q")) 'Set up Address Range 2
    
        For Each account_Address_Cell_2 In account_Address_2_Range.Cells
            If Not IsEmpty(account_Address_Cell_2.Value) Then 'If not empty then load into adjacent column ->
                work_Sheet.Range(work_Sheet.Cells(account_Address_Cell_2.Row, "Q"), work_Sheet.Cells(account_Address_Cell_2.Row, "Q")).Copy
    
                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_2.Row, "R"), final_Sheet.Cells(account_Address_Cell_2.Row, "R")).PasteSpecial xlPasteValues
                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_2.Row, "R"), final_Sheet.Cells(account_Address_Cell_2.Row, "R")).PasteSpecial xlPasteFormats
                
'                Set account_Address_2_Range = final_Sheet.Range(final_Sheet.Cells(2, "R"), final_Sheet.Cells(final_Row_Data, "R")) 'Set up Address Range 2 on Final_Sheet
'
'                For Each account_Address_Cell_2_B In account_Address_2_Range.Cells
'                    For Each invalid_Data In illegal_Characters
'                        account_Address_Cell_2_B = Replace(account_Address_Cell_2_B, invalid_Data, "")
'                        account_Address_Cell_2_B = Replace(account_Address_Cell_2_B, "@", "AT")
'                        account_Address_Cell_2_B = Replace(account_Address_Cell_2_B, "  ", " ")
'                        account_Address_Cell_2_B = Replace(account_Address_Cell_2_B, "&", "AND")
'                    Next invalid_Data
'                Next account_Address_Cell_2_B
'
'                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_2.Row, "R"), final_Sheet.Cells(account_Address_Cell_2.Row, "R")).Replace What:="`,!,@,#,$,%,^", Replacement:="", LookAt:= _
'                    xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_2.Row, "R"), final_Sheet.Cells(account_Address_Cell_2.Row, "R")).Replace What:="@", Replacement:="AT", LookAt:= _
'                    xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_2.Row, "R"), final_Sheet.Cells(account_Address_Cell_2.Row, "R")).Replace What:="&", Replacement:="AND", LookAt:= _
'                    xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_2.Row, "R"), final_Sheet.Cells(account_Address_Cell_2.Row, "R")).Replace What:="  ", Replacement:=" ", LookAt:= _
'                    xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    
                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_2.Row, "S"), final_Sheet.Cells(account_Address_Cell_2.Row, "S")).Value = _
                    "=IF(OR(RIGHT(TRIM(RC[-1]),1)="")"",RIGHT(TRIM(RC[-1]),1)=""."",RIGHT(TRIM(RC[-1]),1)="",""),LEFT(TRIM(RC[-1]),LEN(TRIM(RC[-1]))-1),TRIM(RC[-1]))"
    
                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_2.Row, "S"), final_Sheet.Cells(account_Address_Cell_2.Row, "S")).Copy
                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_2.Row, "R"), final_Sheet.Cells(account_Address_Cell_2.Row, "R")).PasteSpecial xlPasteValues
                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_2.Row, "S"), final_Sheet.Cells(account_Address_Cell_2.Row, "S")).ClearContents
            End If
            If IsEmpty(account_Address_Cell_2.Value) Then  'Checking if cell contains text from Address 1, if empty then load into correct column
                work_Sheet.Range(work_Sheet.Cells(account_Address_Cell_2.Row, "Q"), work_Sheet.Cells(account_Address_Cell_2.Row, "Q")).Copy
    
                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_2.Row, "Q"), final_Sheet.Cells(account_Address_Cell_2.Row, "Q")).PasteSpecial xlPasteValues
                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_2.Row, "Q"), final_Sheet.Cells(account_Address_Cell_2.Row, "Q")).PasteSpecial xlPasteFormats
                
'                Set account_Address_2_Range = final_Sheet.Range(final_Sheet.Cells(2, "Q"), final_Sheet.Cells(final_Row_Data, "Q")) 'Set up Address Range 2 on Final_Sheet
'
'                For Each account_Address_Cell_2_B In account_Address_2_Range.Cells
'                    For Each invalid_Data In illegal_Characters
'                        account_Address_Cell_2_B = Replace(account_Address_Cell_2_B, invalid_Data, "")
'                        account_Address_Cell_2_B = Replace(account_Address_Cell_2_B, "@", "AT")
'                        account_Address_Cell_2_B = Replace(account_Address_Cell_2_B, "  ", " ")
'                        account_Address_Cell_2_B = Replace(account_Address_Cell_2_B, "&", "AND")
'                    Next invalid_Data
'                Next account_Address_Cell_2_B
    
'                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_2.Row, "Q"), final_Sheet.Cells(account_Address_Cell_2.Row, "Q")).Replace What:="`,!,@,#,$,%,^", Replacement:="", LookAt:= _
'                    xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_2.Row, "Q"), final_Sheet.Cells(account_Address_Cell_2.Row, "Q")).Replace What:="@", Replacement:="AT", LookAt:= _
'                    xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_2.Row, "Q"), final_Sheet.Cells(account_Address_Cell_2.Row, "Q")).Replace What:="&", Replacement:="AND", LookAt:= _
'                    xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_2.Row, "Q"), final_Sheet.Cells(account_Address_Cell_2.Row, "Q")).Replace What:="  ", Replacement:=" ", LookAt:= _
'                    xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    
                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_2.Row, "R"), final_Sheet.Cells(account_Address_Cell_2.Row, "R")).Value = _
                    "=IF(OR(RIGHT(TRIM(RC[-1]),1)="")"",RIGHT(TRIM(RC[-1]),1)=""."",RIGHT(TRIM(RC[-1]),1)="",""),LEFT(TRIM(RC[-1]),LEN(TRIM(RC[-1]))-1),TRIM(RC[-1]))"
    
                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_2.Row, "R"), final_Sheet.Cells(account_Address_Cell_2.Row, "R")).Copy
                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_2.Row, "Q"), final_Sheet.Cells(account_Address_Cell_2.Row, "Q")).PasteSpecial xlPasteValues
                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_2.Row, "R"), final_Sheet.Cells(account_Address_Cell_2.Row, "R")).ClearContents
            End If
        Next account_Address_Cell_2
        
'        Set account_Address_2_Range = final_Sheet.Range(final_Sheet.Cells(2, "Q"), final_Sheet.Cells(final_Row_Data, "Q")) 'Set up Address Range 2 on Final_Sheet
'
'        For Each account_Address_Cell_2_B In account_Address_2_Range.Cells
'            For Each invalid_Data In illegal_Characters
'                account_Address_Cell_2_B = Replace(account_Address_Cell_2_B, invalid_Data, "")
'                account_Address_Cell_2_B = Replace(account_Address_Cell_2_B, "@", "AT")
'                account_Address_Cell_2_B = Replace(account_Address_Cell_2_B, "  ", " ")
'                account_Address_Cell_2_B = Replace(account_Address_Cell_2_B, "&", "AND")
'            Next invalid_Data
'        Next account_Address_Cell_2_B
        
        For Each account_Address_Cell_2 In account_Address_2_Range.Cells 'Address 2 Text to Columns
            If Len(account_Address_Cell_2.Value) > 70 Then
                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_2.Row, "Q"), final_Sheet.Cells(account_Address_Cell_2.Row, "Q")).TextToColumns Destination:= _
                    final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_2.Row, "Q"), final_Sheet.Cells(account_Address_Cell_2.Row, "Q")), DataType:=xlFixedWidth, _
                        FieldInfo:=Array(Array(0, 1), Array(70, 1)), TrailingMinusNumbers:=True
            End If
        Next account_Address_Cell_2
    End If

'    For I = 16 To 20 'Columns: P = 16, T = 20
'        final_Row_Data = final_Sheet.Cells(Rows.Count, I).End(xlUp).Row
'        If final_Row_Data = 1 Then
'            'MsgBox (final_Row_Data)
'            final_Row_Data = final_Sheet.Cells(Rows.Count, I - 1).End(xlUp).Row 'Last row of previous non-empty column
'            range_Test = final_Sheet.Range(final_Sheet.Cells(2, "I"), final_Sheet.Cells(final_Row_Data, I - 1)) 'Range of 1st row, Column "P" to last row of last non-empty Column
'            'MsgBox (range_Test)
'            Exit For
'        End If
'    Next I
'    'Try to return final row number in Address Columns P-T if row number = 1 then get last row of previous Column
    
    'final_Row_Data = work_Sheet.Cells(Rows.Count, "Q").End(xlUp).Row 'final_Row for Address 2
    'Set account_Address_1_Range = final_Sheet.Range(final_Sheet.Cells(2, "P"), final_Sheet.Cells(final_Row_Data, "T")) 'Set up Address Range 1 on Final_Sheet

'    For Each account_Address_Cell_1 In account_Address_1_Range.Cells
'        For Each invalid_Data In illegal_Characters
'            account_Address_Cell_1 = Replace(account_Address_Cell_1, invalid_Data, "")
'            account_Address_Cell_1 = Replace(account_Address_Cell_1, "@", "AT")
'            account_Address_Cell_1 = Replace(account_Address_Cell_1, "  ", " ")
'            account_Address_Cell_1 = Replace(account_Address_Cell_1, "&", "AND")
'        Next invalid_Data
'    Next account_Address_Cell_1
    
        
    final_Row_Data = work_Sheet.Cells(Rows.Count, "R").End(xlUp).Row 'final_Row for Address 3
    
    If final_Row_Data <> 1 Then 'Address 3,final_Row_Data cannot be 1
        Set account_Address_3_Range = final_Sheet.Range(final_Sheet.Cells(2, "R"), final_Sheet.Cells(final_Row_Data, "R")) 'Set up Address Range 3
        
        For Each account_Address_Cell_3 In account_Address_3_Range.Cells
            If Not IsEmpty(account_Address_Cell_3.Value) Then 'If not empty then load into adjacent column ->
                work_Sheet.Range(work_Sheet.Cells(account_Address_Cell_3.Row, "R"), work_Sheet.Cells(account_Address_Cell_3.Row, "R")).Copy

'                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_3.Row, "S"), final_Sheet.Cells(account_Address_Cell_3.Row, "S")).PasteSpecial xlPasteValues
'                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_3.Row, "S"), final_Sheet.Cells(account_Address_Cell_3.Row, "S")).PasteSpecial xlPasteFormats
'
'                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_3.Row, "S"), final_Sheet.Cells(account_Address_Cell_3.Row, "S")).Replace What:="`,!,@,#,$,%,^", Replacement:="", LookAt:= _
'                    xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_3.Row, "S"), final_Sheet.Cells(account_Address_Cell_3.Row, "S")).Replace What:="@", Replacement:="AT", LookAt:= _
'                    xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_3.Row, "S"), final_Sheet.Cells(account_Address_Cell_3.Row, "S")).Replace What:="&", Replacement:="AND", LookAt:= _
'                    xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_3.Row, "S"), final_Sheet.Cells(account_Address_Cell_3.Row, "S")).Replace What:="  ", Replacement:=" ", LookAt:= _
'                    xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_3.Row, "T"), final_Sheet.Cells(account_Address_Cell_3.Row, "T")).Value = _
                    "=IF(OR(RIGHT(TRIM(RC[-1]),1)="")"",RIGHT(TRIM(RC[-1]),1)=""."",RIGHT(TRIM(RC[-1]),1)="",""),LEFT(TRIM(RC[-1]),LEN(TRIM(RC[-1]))-1),TRIM(RC[-1]))"

                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_3.Row, "T"), final_Sheet.Cells(account_Address_Cell_3.Row, "T")).Copy
                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_3.Row, "S"), final_Sheet.Cells(account_Address_Cell_3.Row, "S")).PasteSpecial xlPasteValues
                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_3.Row, "T"), final_Sheet.Cells(account_Address_Cell_3.Row, "T")).ClearContents
            End If
            If IsEmpty(account_Address_Cell_3.Value) Then  'Checking if cell contains text from Address 2, if empty then load into correct column
                work_Sheet.Range(work_Sheet.Cells(account_Address_Cell_3.Row, "R"), work_Sheet.Cells(account_Address_Cell_3.Row, "R")).Copy

'                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_3.Row, "R"), final_Sheet.Cells(account_Address_Cell_3.Row, "R")).PasteSpecial xlPasteValues
'                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_3.Row, "R"), final_Sheet.Cells(account_Address_Cell_3.Row, "R")).PasteSpecial xlPasteFormats
'
'                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_3.Row, "R"), final_Sheet.Cells(account_Address_Cell_3.Row, "R")).Replace What:="`,!,@,#,$,%,^", Replacement:="", LookAt:= _
'                    xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_3.Row, "R"), final_Sheet.Cells(account_Address_Cell_3.Row, "R")).Replace What:="@", Replacement:="AT", LookAt:= _
'                    xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_3.Row, "R"), final_Sheet.Cells(account_Address_Cell_3.Row, "R")).Replace What:="&", Replacement:="AND", LookAt:= _
'                    xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_3.Row, "R"), final_Sheet.Cells(account_Address_Cell_3.Row, "R")).Replace What:="  ", Replacement:=" ", LookAt:= _
'                    xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_3.Row, "S"), final_Sheet.Cells(account_Address_Cell_3.Row, "S")).Value = _
                    "=IF(OR(RIGHT(TRIM(RC[-1]),1)="")"",RIGHT(TRIM(RC[-1]),1)=""."",RIGHT(TRIM(RC[-1]),1)="",""),LEFT(TRIM(RC[-1]),LEN(TRIM(RC[-1]))-1),TRIM(RC[-1]))"

                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_3.Row, "S"), final_Sheet.Cells(account_Address_Cell_3.Row, "S")).Copy
                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_3.Row, "R"), final_Sheet.Cells(account_Address_Cell_3.Row, "R")).PasteSpecial xlPasteValues
                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_3.Row, "S"), final_Sheet.Cells(account_Address_Cell_3.Row, "S")).ClearContents
            End If
        Next account_Address_Cell_3
        
        For Each account_Address_Cell_3 In account_Address_3_Range.Cells 'Address 3 Text to Columns
            If Len(account_Address_Cell_3.Value) > 70 Then
                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_3.Row, "R"), final_Sheet.Cells(account_Address_Cell_3.Row, "R")).TextToColumns Destination:= _
                    final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_3.Row, "R"), final_Sheet.Cells(account_Address_Cell_3.Row, "R")), DataType:=xlFixedWidth, _
                        FieldInfo:=Array(Array(0, 1), Array(70, 1)), TrailingMinusNumbers:=True
            End If
        Next account_Address_Cell_3
    End If
    
    'final_Row_Data = work_Sheet.Cells(Rows.Count, "Q").End(xlUp).Row 'final_Row for Address 2
    'Set account_Address_1_Range = final_Sheet.Range(final_Sheet.Cells(2, "P"), final_Sheet.Cells(final_Row_Data, "T")) 'Set up Address Range 1 on Final_Sheet

'    For Each account_Address_Cell_1 In account_Address_1_Range.Cells
'        For Each invalid_Data In illegal_Characters
'            account_Address_Cell_1 = Replace(account_Address_Cell_1, invalid_Data, "")
'            account_Address_Cell_1 = Replace(account_Address_Cell_1, "@", "AT")
'            account_Address_Cell_1 = Replace(account_Address_Cell_1, "  ", " ")
'            account_Address_Cell_1 = Replace(account_Address_Cell_1, "&", "AND")
'        Next invalid_Data
'    Next account_Address_Cell_1
'
'    For Each account_Address_Cell_Final In column_A_To_AA.Cells
'        For Each invalid_Data In illegal_Characters
'            account_Address_Cell_Final = Replace(account_Address_Cell_Final, invalid_Data, "")
'            account_Address_Cell_Final = Replace(account_Address_Cell_Final, "@", "AT")
'            account_Address_Cell_Final = Replace(account_Address_Cell_Final, "  ", " ")
'            account_Address_Cell_Final = Replace(account_Address_Cell_Final, "&", "AND")
'        Next invalid_Data
'    Next account_Address_Cell_Final

'        With final_Sheet.Cells(final_Row_Data, "P")
'                .HorizontalAlignment = xlLeft
'                .VerticalAlignment = xlBottom
'                .WrapText = False
'                .Orientation = 0
'                .AddIndent = False
'                .IndentLevel = 0
'                .ShrinkToFit = False
'                .ReadingOrder = xlContext
'                .MergeCells = False
'            End With
    End If
    final_Row_Data = work_Sheet.Cells(Rows.Count, "AA").End(xlUp).Row
    Set column_A_To_AA = final_Sheet.Range(final_Sheet.Cells(2, "A"), final_Sheet.Cells(final_Row_Data, "AA"))
    
    column_A_To_AA.Replace What:="`", Replacement:="", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

    column_A_To_AA.Replace What:="!", Replacement:="", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        
    column_A_To_AA.Replace What:="@", Replacement:="AT", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        
    column_A_To_AA.Replace What:="#", Replacement:="", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        
    column_A_To_AA.Replace What:="$", Replacement:="", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        
    column_A_To_AA.Replace What:="%", Replacement:="", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

    column_A_To_AA.Replace What:="^", Replacement:="AT", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

    column_A_To_AA.Replace What:="&", Replacement:="AND", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

    column_A_To_AA.Replace What:="  ", Replacement:=" ", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        
    column_A_To_AA.Replace What:="Ä", Replacement:="A", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        
    column_A_To_AA.Replace What:="Ê", Replacement:="E", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        
    column_A_To_AA.Replace What:="Ï", Replacement:="I", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        
    column_A_To_AA.Replace What:="Ö", Replacement:="O", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        
    column_A_To_AA.Replace What:="Ü", Replacement:="U", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        
    column_A_To_AA.Replace What:="Ÿ", Replacement:="Y", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        
    With column_A_To_AA.Cells
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
End Sub
Private Sub Street_Name()
    'Dim Final_Row_Street_Name
    Dim work_Book As Workbook
    Dim work_Sheet As Worksheet
    'Dim calculate_Sheet As Worksheet
    Dim final_Sheet As Worksheet
    Set work_Book = ThisWorkbook
    Set work_Sheet = work_Book.Worksheets("Sheet1")
    'Set calculate_Sheet = work_Book.Worksheets("Calculate_Sheet")
    Set final_Sheet = work_Book.Worksheets("Final_Sheet")
    final_Row_Data = Cells(Rows.Count, "U").End(xlUp).Row
    
    If final_Row_Data <> 1 Then
        work_Sheet.Range(work_Sheet.Cells(final_Row_Data, "U"), work_Sheet.Cells(2, "U")).Copy
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "U"), final_Sheet.Cells(2, "U")).PasteSpecial xlPasteValues
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "U"), final_Sheet.Cells(2, "U")).PasteSpecial xlPasteFormats
        
'        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "U"), final_Sheet.Cells(2, "U")).Replace What:="`,!,@,#,$,%,^", Replacement:="", LookAt:= _
'            xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "U"), final_Sheet.Cells(2, "U")).Replace What:="@", Replacement:="AT", LookAt:= _
'            xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "U"), final_Sheet.Cells(2, "U")).Replace What:="&", Replacement:="AND", LookAt:= _
'            xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "U"), final_Sheet.Cells(2, "U")).Replace What:="  ", Replacement:=" ", LookAt:= _
'            xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
            
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "V"), final_Sheet.Cells(2, "V")).Value = _
            "=IF(OR(RIGHT(TRIM(RC[-1]),1)="")"",RIGHT(TRIM(RC[-1]),1)=""."",RIGHT(TRIM(RC[-1]),1)="",""),LEFT(TRIM(RC[-1]),LEN(TRIM(RC[-1]))-1),TRIM(RC[-1]))"
        
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "V"), final_Sheet.Cells(2, "V")).Copy
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "U"), final_Sheet.Cells(2, "U")).PasteSpecial xlPasteValues
        'final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "U"), final_Sheet.Cells(2, "U")).PasteSpecial xlPasteFormats
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "V"), final_Sheet.Cells(2, "V")).ClearContents
        
'        With final_Sheet.Cells(final_Row_Data, "U")
'            .HorizontalAlignment = xlLeft
'            .VerticalAlignment = xlBottom
'            .WrapText = False
'            .Orientation = 0
'            .AddIndent = False
'            .IndentLevel = 0
'            .ShrinkToFit = False
'            .ReadingOrder = xlContext
'            .MergeCells = False
'        End With
    End If
End Sub
Private Sub Building_No()
    'Dim Final_Row_Building_No
    Dim work_Book As Workbook
    Dim work_Sheet As Worksheet
    'Dim calculate_Sheet As Worksheet
    Dim final_Sheet As Worksheet
    Set work_Book = ThisWorkbook
    Set work_Sheet = work_Book.Worksheets("Sheet1")
    'Set calculate_Sheet = work_Book.Worksheets("Calculate_Sheet")
    Set final_Sheet = work_Book.Worksheets("Final_Sheet")
    final_Row_Data = work_Sheet.Cells(Rows.Count, "V").End(xlUp).Row
    
    If final_Row_Data <> 1 Then
        work_Sheet.Range(work_Sheet.Cells(final_Row_Data, "V"), work_Sheet.Cells(2, "V")).Copy
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "V"), final_Sheet.Cells(2, "V")).PasteSpecial xlPasteValues
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "V"), final_Sheet.Cells(2, "V")).PasteSpecial xlPasteFormats ' ERROR WITH COPY !!!!!!!!!!!!!!!!!!!!!!!
        
'        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "V"), final_Sheet.Cells(2, "V")).Replace What:="`,!,@,#,$,%,^", Replacement:="", LookAt:= _
'            xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "V"), final_Sheet.Cells(2, "V")).Replace What:="@", Replacement:="AT", LookAt:= _
'            xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "V"), final_Sheet.Cells(2, "V")).Replace What:="&", Replacement:="AND", LookAt:= _
'            xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "V"), final_Sheet.Cells(2, "V")).Replace What:="  ", Replacement:=" ", LookAt:= _
'            xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
            
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "W"), final_Sheet.Cells(2, "W")).Value = _
            "=IF(OR(RIGHT(TRIM(RC[-1]),1)="")"",RIGHT(TRIM(RC[-1]),1)=""."",RIGHT(TRIM(RC[-1]),1)="",""),LEFT(TRIM(RC[-1]),LEN(TRIM(RC[-1]))-1),TRIM(RC[-1]))"
        
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "W"), final_Sheet.Cells(2, "W")).Copy
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "V"), final_Sheet.Cells(2, "V")).PasteSpecial xlPasteValues
        'final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "V"), final_Sheet.Cells(2, "V")).PasteSpecial xlPasteFormats
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "W"), final_Sheet.Cells(2, "W")).ClearContents
        
'        With final_Sheet.Cells(final_Row_Data, "V")
'            .HorizontalAlignment = xlLeft
'            .VerticalAlignment = xlBottom
'            .WrapText = False
'            .Orientation = 0
'            .AddIndent = False
'            .IndentLevel = 0
'            .ShrinkToFit = False
'            .ReadingOrder = xlContext
'            .MergeCells = False
'        End With
    End If

    
End Sub
Private Sub PO_Box()
    'Dim Final_Row_PO_Box
    Dim work_Book As Workbook
    Dim work_Sheet As Worksheet
    'Dim calculate_Sheet As Worksheet
    Dim final_Sheet As Worksheet
    Set work_Book = ThisWorkbook
    Set work_Sheet = work_Book.Worksheets("Sheet1")
    'Set calculate_Sheet = work_Book.Worksheets("Calculate_Sheet")
    Set final_Sheet = work_Book.Worksheets("Final_Sheet")
    final_Row_Data = work_Sheet.Cells(Rows.Count, "W").End(xlUp).Row
    
    If final_Row_Data <> 1 Then
        work_Sheet.Range(work_Sheet.Cells(final_Row_Data, "W"), work_Sheet.Cells(2, "W")).Copy
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "W"), final_Sheet.Cells(2, "W")).PasteSpecial xlPasteValues
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "W"), final_Sheet.Cells(2, "W")).PasteSpecial xlPasteFormats
        
'        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "W"), final_Sheet.Cells(2, "W")).Replace What:="`,!,@,#,$,%,^", Replacement:="", LookAt:= _
'            xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "W"), final_Sheet.Cells(2, "W")).Replace What:="@", Replacement:="AT", LookAt:= _
'            xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "W"), final_Sheet.Cells(2, "W")).Replace What:="&", Replacement:="AND", LookAt:= _
'            xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "W"), final_Sheet.Cells(2, "W")).Replace What:="  ", Replacement:=" ", LookAt:= _
'            xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
            
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "X"), final_Sheet.Cells(2, "X")).Value = _
            "=IF(OR(RIGHT(TRIM(RC[-1]),1)="")"",RIGHT(TRIM(RC[-1]),1)=""."",RIGHT(TRIM(RC[-1]),1)="",""),LEFT(TRIM(RC[-1]),LEN(TRIM(RC[-1]))-1),TRIM(RC[-1]))"
        
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "X"), final_Sheet.Cells(2, "X")).Copy
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "W"), final_Sheet.Cells(2, "W")).PasteSpecial xlPasteValues
        'final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "W"), final_Sheet.Cells(2, "W")).PasteSpecial xlPasteFormats
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "X"), final_Sheet.Cells(2, "X")).ClearContents
        
'        With final_Sheet.Cells(final_Row_Data, "W")
'            .HorizontalAlignment = xlLeft
'            .VerticalAlignment = xlBottom
'            .WrapText = False
'            .Orientation = 0
'            .AddIndent = False
'            .IndentLevel = 0
'            .ShrinkToFit = False
'            .ReadingOrder = xlContext
'            .MergeCells = False
'        End With
    End If
End Sub
Private Sub Postal_Code()
    'Dim Final_Row_Postal_Code
    Dim work_Book As Workbook
    Dim work_Sheet As Worksheet
    'Dim calculate_Sheet As Worksheet
    Dim final_Sheet As Worksheet
    Set work_Book = ThisWorkbook
    Set work_Sheet = work_Book.Worksheets("Sheet1")
    'Set calculate_Sheet = work_Book.Worksheets("Calculate_Sheet")
    Set final_Sheet = work_Book.Worksheets("Final_Sheet")
    final_Row_Data = work_Sheet.Cells(Rows.Count, "X").End(xlUp).Row
    
    If final_Row_Data <> 1 Then
        work_Sheet.Range(work_Sheet.Cells(final_Row_Data, "X"), work_Sheet.Cells(2, "X")).Copy
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "X"), final_Sheet.Cells(2, "X")).PasteSpecial xlPasteValues
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "X"), final_Sheet.Cells(2, "X")).PasteSpecial xlPasteFormats
        
'        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "X"), final_Sheet.Cells(2, "X")).Replace What:="`,!,@,#,$,%,^", Replacement:="", LookAt:= _
'            xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "X"), final_Sheet.Cells(2, "X")).Replace What:="@", Replacement:="AT", LookAt:= _
'            xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "X"), final_Sheet.Cells(2, "X")).Replace What:="&", Replacement:="AND", LookAt:= _
'            xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "X"), final_Sheet.Cells(2, "X")).Replace What:="  ", Replacement:=" ", LookAt:= _
'            xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
            
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "Y"), final_Sheet.Cells(2, "Y")).Value = _
            "=IF(OR(RIGHT(TRIM(RC[-1]),1)="")"",RIGHT(TRIM(RC[-1]),1)=""."",RIGHT(TRIM(RC[-1]),1)="",""),LEFT(TRIM(RC[-1]),LEN(TRIM(RC[-1]))-1),TRIM(RC[-1]))"
        
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "Y"), final_Sheet.Cells(2, "Y")).Copy
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "X"), final_Sheet.Cells(2, "X")).PasteSpecial xlPasteValues
        'final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "X"), final_Sheet.Cells(2, "X")).PasteSpecial xlPasteFormats
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "Y"), final_Sheet.Cells(2, "Y")).ClearContents
        
        
'        With final_Sheet.Cells(final_Row_Data, "X")
'            .HorizontalAlignment = xlLeft
'            .VerticalAlignment = xlBottom
'            .WrapText = False
'            .Orientation = 0
'            .AddIndent = False
'            .IndentLevel = 0
'            .ShrinkToFit = False
'            .ReadingOrder = xlContext
'            .MergeCells = False
'        End With
    End If
End Sub
Private Sub Town()
    'Dim Final_Row_Town
    Dim work_Book As Workbook
    Dim work_Sheet As Worksheet
    'Dim calculate_Sheet As Worksheet
    Dim final_Sheet As Worksheet
    Set work_Book = ThisWorkbook
    Set work_Sheet = work_Book.Worksheets("Sheet1")
    'Set calculate_Sheet = work_Book.Worksheets("Calculate_Sheet")
    Set final_Sheet = work_Book.Worksheets("Final_Sheet")
    final_Row_Data = work_Sheet.Cells(Rows.Count, "Y").End(xlUp).Row
    
    
    If final_Row_Data <> 1 Then
        work_Sheet.Range(work_Sheet.Cells(final_Row_Data, "Y"), work_Sheet.Cells(2, "Y")).Copy
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "Y"), final_Sheet.Cells(2, "Y")).PasteSpecial xlPasteValues
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "Y"), final_Sheet.Cells(2, "Y")).PasteSpecial xlPasteFormats
        
'        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "Y"), final_Sheet.Cells(2, "Y")).Replace What:="`,!,@,#,$,%,^", Replacement:="", LookAt:= _
'            xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "Y"), final_Sheet.Cells(2, "Y")).Replace What:="@", Replacement:="AT", LookAt:= _
'            xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "Y"), final_Sheet.Cells(2, "Y")).Replace What:="&", Replacement:="AND", LookAt:= _
'            xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "Y"), final_Sheet.Cells(2, "Y")).Replace What:="  ", Replacement:=" ", LookAt:= _
'            xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
            
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "Z"), final_Sheet.Cells(2, "Z")).Value = _
            "=IF(OR(RIGHT(TRIM(RC[-1]),1)="")"",RIGHT(TRIM(RC[-1]),1)=""."",RIGHT(TRIM(RC[-1]),1)="",""),LEFT(TRIM(RC[-1]),LEN(TRIM(RC[-1]))-1),TRIM(RC[-1]))"
        
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "Z"), final_Sheet.Cells(2, "Z")).Copy
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "Y"), final_Sheet.Cells(2, "Y")).PasteSpecial xlPasteValues
        'final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "Y"), final_Sheet.Cells(2, "Y")).PasteSpecial xlPasteFormats
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "Z"), final_Sheet.Cells(2, "Z")).ClearContents
        
'        With final_Sheet.Cells(final_Row_Data, "Y")
'            .HorizontalAlignment = xlLeft
'            .VerticalAlignment = xlBottom
'            .WrapText = False
'            .Orientation = 0
'            .AddIndent = False
'            .IndentLevel = 0
'            .ShrinkToFit = False
'            .ReadingOrder = xlContext
'            .MergeCells = False
'        End With
    End If
    
End Sub
Private Sub Province()
    'Dim Final_Row_Province
    Dim work_Book As Workbook
    Dim work_Sheet As Worksheet
    'Dim calculate_Sheet As Worksheet
    Dim final_Sheet As Worksheet
    Set work_Book = ThisWorkbook
    Set work_Sheet = work_Book.Worksheets("Sheet1")
    'Set calculate_Sheet = work_Book.Worksheets("Calculate_Sheet")
    Set final_Sheet = work_Book.Worksheets("Final_Sheet")
    final_Row_Data = work_Sheet.Cells(Rows.Count, "Z").End(xlUp).Row
    
    work_Sheet.Range(work_Sheet.Cells(final_Row_Data, "Z"), work_Sheet.Cells(2, "Z")).Copy
    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "Z"), final_Sheet.Cells(2, "Z")).PasteSpecial xlPasteValues
    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "Z"), final_Sheet.Cells(2, "Z")).PasteSpecial xlPasteFormats
    
'    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "Z"), final_Sheet.Cells(2, "Z")).Replace What:="`,!,@,#,$,%,^", Replacement:="", LookAt:= _
'        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "Z"), final_Sheet.Cells(2, "Z")).Replace What:="@", Replacement:="AT", LookAt:= _
'        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "Z"), final_Sheet.Cells(2, "Z")).Replace What:="&", Replacement:="AND", LookAt:= _
'        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "Z"), final_Sheet.Cells(2, "Z")).Replace What:="  ", Replacement:=" ", LookAt:= _
'        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        
    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "AA"), final_Sheet.Cells(2, "AA")).Value = _
        "=IF(OR(RIGHT(TRIM(RC[-1]),1)="")"",RIGHT(TRIM(RC[-1]),1)=""."",RIGHT(TRIM(RC[-1]),1)="",""),LEFT(TRIM(RC[-1]),LEN(TRIM(RC[-1]))-1),TRIM(RC[-1]))"
    
    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "AA"), final_Sheet.Cells(2, "AA")).Copy
    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "Z"), final_Sheet.Cells(2, "Z")).PasteSpecial xlPasteValues
    'final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "Z"), final_Sheet.Cells(2, "Z")).PasteSpecial xlPasteFormats
    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "AA"), final_Sheet.Cells(2, "AA")).ClearContents
    
'    For Counter = 3 To final_Row_Data
'        calculate_Sheet.Cells(Counter, 1).Value = work_Sheet.Cells(Counter, "Z").Value
'        calculate_Sheet.Cells(Counter, 2).Value = calculate_Sheet.Cells(Counter, 1).Value
'
'        calculate_Sheet.Cells(Counter, 2).Replace What:="`", Replacement:="", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        calculate_Sheet.Cells(Counter, 2).Replace What:="!", Replacement:="", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        calculate_Sheet.Cells(Counter, 2).Replace What:="@", Replacement:="AT", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        calculate_Sheet.Cells(Counter, 2).Replace What:="#", Replacement:="", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        calculate_Sheet.Cells(Counter, 2).Replace What:="$", Replacement:="", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        calculate_Sheet.Cells(Counter, 2).Replace What:="%", Replacement:="", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        calculate_Sheet.Cells(Counter, 2).Replace What:="^", Replacement:="", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        calculate_Sheet.Cells(Counter, 2).Replace What:="&", Replacement:="AND", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        calculate_Sheet.Cells(Counter, 2).Replace What:="  ", Replacement:=" ", LookAt:=xlPart, SearchOrder:= _
'            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        calculate_Sheet.Cells(Counter, 3).Value = "=IF(OR(RIGHT(TRIM(RC[-1]),1)="")"",RIGHT(TRIM(RC[-1]),1)=""."",RIGHT(TRIM(RC[-1]),1)="",""),LEFT(TRIM(RC[-1]),LEN(TRIM(RC[-1]))-1),TRIM(RC[-1]))"
'
'        calculate_Sheet.Cells(Counter, 3).Copy
'        final_Sheet.Cells(Counter - 1, "Z").PasteSpecial xlPasteValues
'
'        With final_Sheet.Cells(final_Row_Data, "Z")
'            .HorizontalAlignment = xlLeft
'            .VerticalAlignment = xlBottom
'            .WrapText = False
'            .Orientation = 0
'            .AddIndent = False
'            .IndentLevel = 0
'            .ShrinkToFit = False
'            .ReadingOrder = xlContext
'            .MergeCells = False
'        End With
'        'final_Sheet.Cells(Counter - 1, "Z").Value = calculate_Sheet.Cells(Counter, 4).Value
'    Next Counter
End Sub
Private Sub Country_Code()
    'Dim Final_Row_Country_Code
    Dim work_Book As Workbook
    Dim work_Sheet As Worksheet
    'Dim calculate_Sheet As Worksheet
    Dim final_Sheet As Worksheet
    Set work_Book = ThisWorkbook
    Set work_Sheet = work_Book.Worksheets("Sheet1")
    'Set calculate_Sheet = work_Book.Worksheets("Calculate_Sheet")
    Set final_Sheet = work_Book.Worksheets("Final_Sheet")
    final_Row_Data = work_Sheet.Cells(Rows.Count, "AA").End(xlUp).Row
    
    If final_Row_Data <> 1 Then
        work_Sheet.Range(work_Sheet.Cells(final_Row_Data, "AA"), work_Sheet.Cells(2, "AA")).Copy
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "AA"), final_Sheet.Cells(2, "AA")).PasteSpecial xlPasteValues
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "AA"), final_Sheet.Cells(2, "AA")).PasteSpecial xlPasteFormats
        
'        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "AA"), final_Sheet.Cells(2, "AA")).Replace What:="`,!,@,#,$,%,^", Replacement:="", LookAt:= _
'            xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "AA"), final_Sheet.Cells(2, "AA")).Replace What:="@", Replacement:="AT", LookAt:= _
'            xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "AA"), final_Sheet.Cells(2, "AA")).Replace What:="&", Replacement:="AND", LookAt:= _
'            xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "AA"), final_Sheet.Cells(2, "AA")).Replace What:="  ", Replacement:=" ", LookAt:= _
'            xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
            
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "AB"), final_Sheet.Cells(2, "AB")).Value = _
            "=IF(OR(RIGHT(TRIM(RC[-1]),1)="")"",RIGHT(TRIM(RC[-1]),1)=""."",RIGHT(TRIM(RC[-1]),1)="",""),LEFT(TRIM(RC[-1]),LEN(TRIM(RC[-1]))-1),TRIM(RC[-1]))"
        
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "AB"), final_Sheet.Cells(2, "AB")).Copy
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "AA"), final_Sheet.Cells(2, "AA")).PasteSpecial xlPasteValues
        'final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "AA"), final_Sheet.Cells(2, "AA")).PasteSpecial xlPasteFormats
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "AB"), final_Sheet.Cells(2, "AB")).ClearContents
        
'            With final_Sheet.Cells(final_Row_Data, "AA")
'                .HorizontalAlignment = xlLeft
'                .VerticalAlignment = xlBottom
'                .WrapText = False
'                .Orientation = 0
'                .AddIndent = False
'                .IndentLevel = 0
'                .ShrinkToFit = False
'                .ReadingOrder = xlContext
'                .MergeCells = False
'            End With
    End If
End Sub
Private Sub Account_Length() 'Check if Account Number Length is correct for User_Input CID
    'Dim Final_Row_Account_Num As Long
    Dim account_Num_Range As Range
    Dim account_Num_Cell As Range
    Dim work_Book As Workbook
    Dim work_Sheet As Worksheet
    'Dim calculate_Sheet As Worksheet
    Dim final_Sheet As Worksheet
    Set work_Book = ThisWorkbook
    Set work_Sheet = work_Book.Worksheets("Sheet1")
    'Set calculate_Sheet = work_Book.Worksheets("Calculate_Sheet")
    Set final_Sheet = work_Book.Worksheets("Final_Sheet")
    Dim Account_Num_Length As String
    Dim CID As String
    'Dim Incorrect_Account_Length As Boolean
    CID = InputBox("Enter CID")
    final_Row_Data = work_Sheet.Cells(work_Sheet.Rows.Count, "B").End(xlUp).Row
    Set account_Num_Range = work_Sheet.Range(work_Sheet.Cells(2, 2), work_Sheet.Cells(final_Row_Data, 2))
    
    For Each account_Num_Cell In account_Num_Range.Cells
        If Len(account_Num_Cell.Value) <> 5 And CID = "55P" Then
            incorrect_Account_Length = True
            work_Sheet.Cells(account_Num_Cell.Address, "B").Interior.ColorIndex = 35
        End If
        If Len(account_Num_Cell.Value) <> 11 And CID = "5DU" Then
            incorrect_Account_Length = True
            work_Sheet.Cells(account_Num_Cell.Row, "B").Interior.ColorIndex = 35
        End If
    Next account_Num_Cell
    
'    If Len(work_Sheet.Range(Cells(3, "B"), Cells(final_Row_Data, "B")).Value) <> 5 And CID = "55P" Then
'        incorrect_Account_Length = True
'        work_Sheet.Cells(Counter, 2).Interior.ColorIndex = 35
'    End If
'
'    For Counter = 3 To final_Row_Data
'        If Len(work_Sheet.Cells(Counter, 2).Value) <> 5 And CID = "55P" Then
'            incorrect_Account_Length = True
'            work_Sheet.Cells(Counter, 2).Interior.ColorIndex = 35
'        End If
'        If Len(work_Sheet.Cells(Counter, 2).Value) <> 11 And CID = "5DU" Then
'            incorrect_Account_Length = True
'            work_Sheet.Cells(Counter, 2).Interior.ColorIndex = 35
'        End If
'    Next Counter
'    If incorrect_Account_Length Then
'        MsgBox ("Cells with Incorrect Account Length have been hightlighted and sorted for you")
'        'incorrect_Account_Length = False
'        work_Sheet.Range("B:B").AutoFilter Field:=1, Criteria1:=RGB(204, 255 _
'        , 204), Operator:=xlFilterCellColor
'    End If
End Sub
Private Sub Create_Final_Sheet() 'Create a Calculate_Sheet and Final_Sheet if none exist
Attribute Create_Final_Sheet.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim work_Book As Workbook
    Set work_Book = ThisWorkbook
    Dim sheet_Calculate As String
    Dim sheet_Final As String
    'sheet_Calculate = "Calculate_Sheet"
    sheet_Final = "Final_Sheet"
    'Dim sheet_Calculate_Exists As Boolean
    Dim sheet_Final_Exists As Boolean

    For Each work_Sheet In work_Book.Worksheets
        'If work_Sheet.Name = "Calculate_Sheet" Then
        '    sheet_Calculate_Exists = True
        'End If
        If work_Sheet.Name = "Final_Sheet" Then
            sheet_Final_Exists = True
        End If
    Next work_Sheet
    
    'If sheet_Calculate_Exists = False Then
    '    work_Book.Worksheets.Add After:=Worksheets("Sheet1")
    '    Worksheets(2).Name = "Calculate_Sheet"
        'ActiveSheet.Name = "Calculate_Sheet"
        'work_Sheet.Select
    'End If
    
    If sheet_Final_Exists = False Then
        work_Book.Worksheets.Add After:=Worksheets("Sheet1").Name = "Final_Sheet"
        'Worksheets(2).Name = "Final_Sheet"
        'ActiveSheet.Name = "Final_Sheet"
        'work_Sheet.Select
    End If
End Sub
