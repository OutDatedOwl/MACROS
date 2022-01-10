Attribute VB_Name = "Module2"
    Global work_Sheet As Worksheet
    Global work_Book As ThisWorkbook
    Option Compare Text
Sub Main() 'Main BO_Adder
Attribute Main.VB_ProcData.VB_Invoke_Func = " \n14"
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    'Call check_Columns_For_Text
    Create_Calculation_Sheet
    Check_CID
    Account_Number
    Market_ID
    Entity_Type_Code
    BO_Name
    URN_Type
    URN_Number
    Address_Type
    Address_1
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
    
    Worksheets("Sheet1").Rows("1:1").Copy
    Worksheets("Final_Sheet").Rows("1:1").PasteSpecial xlPasteValues
    Worksheets("Final_Sheet").Rows("1:1").PasteSpecial xlFormats
    Application.ScreenUpdating = True
End Sub
Private Sub check_Columns_For_Text() 'Check if columns are empty
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
            If Worksheets("Sheet1").Cells(Counter_B, Counter).Value <> "" Then
                empty_Column = False
                'work_Sheet.Cells(Counter_B, Counter).Interior.ColorIndex = 5
                'MsgBox ("Row " & Counter_B & " Column " & Counter)
            End If
        Next Counter_B
        'MsgBox (Counter)
        'MsgBox "Column " & Counter & " Empty: " & IsEmpty(Range("D3:D"))
        If empty_Column = True Then
            'MsgBox ("Column " & Counter & " is Empty")
            'MsgBox (Counter)
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
    Dim Final_Row_Client_ID
    Dim work_Book As ThisWorkbook
    Dim work_Sheet As Worksheet
    Dim calculate_Sheet As Worksheet
    Dim final_Sheet As Worksheet
    Set work_Book = ThisWorkbook
    Set work_Sheet = work_Book.Worksheets("Sheet1")
    Set calculate_Sheet = work_Book.Worksheets("Calculate_Sheet")
    Set final_Sheet = work_Book.Worksheets("Final_Sheet")
    Final_Row_Client_ID = work_Sheet.Cells(Rows.Count, "B").End(xlUp).Row
    
    For Counter = 3 To Final_Row_Client_ID
        calculate_Sheet.Cells(Counter, 1).Value = work_Sheet.Cells(Counter, "B").Value
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
        calculate_Sheet.Cells(Counter, 4).Copy
        Worksheets("Final_Sheet").Cells(Counter - 1, "B").PasteSpecial xlPasteValues
        Worksheets("Final_Sheet").Cells(Counter - 1, "B").PasteSpecial xlFormats
    Next Counter
    Account_Length
End Sub
Private Sub Market_ID() 'Fill in Markets Column
    Dim Final_Row_Market_ID
    Dim work_Book As Workbook
    Dim work_Sheet As Worksheet
    Dim calculate_Sheet As Worksheet
    Dim final_Sheet As Worksheet
    Set work_Book = ThisWorkbook
    Set work_Sheet = work_Book.Worksheets("Sheet1")
    Set calculate_Sheet = work_Book.Worksheets("Calculate_Sheet")
    Set final_Sheet = work_Book.Worksheets("Final_Sheet")
    Final_Row_Market_ID = work_Sheet.Cells(Rows.Count, 1).End(xlUp).Row
    
    final_Sheet.Select
    final_Sheet.Range(final_Sheet.Cells(Final_Row_Market_ID - 1, "D"), final_Sheet.Cells(2, "D")).Value = "***"
    work_Sheet.Select
End Sub
Private Sub Entity_Type_Code()
    Dim Final_Row_Entity_Code_ID
    Dim work_Book As Workbook
    Dim work_Sheet As Worksheet
    Dim calculate_Sheet As Worksheet
    Dim final_Sheet As Worksheet
    Set work_Book = ThisWorkbook
    Set work_Sheet = work_Book.Worksheets("Sheet1")
    Set calculate_Sheet = work_Book.Worksheets("Calculate_Sheet")
    Set final_Sheet = work_Book.Worksheets("Final_Sheet")
    Final_Row_Entity_Code_ID = Cells(Rows.Count, 1).End(xlUp).Row
    
    For Counter = 3 To Final_Row_Entity_Code_ID
        calculate_Sheet.Cells(Counter, 1).Value = work_Sheet.Cells(Counter, "E").Value
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
            
        With Sheets("Calculate_Sheet").Cells(Counter, 4)
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
        Worksheets("Final_Sheet").Cells(Counter - 1, "E").Value = calculate_Sheet.Cells(Counter, 4).Value
    Next Counter
End Sub
Private Sub Check_CID() 'Checking for CID
    Dim Final_Row_Client_ID
    Dim work_Book As Workbook
    Dim work_Sheet As Worksheet
    Dim calculate_Sheet As Worksheet
    Dim final_Sheet As Worksheet
    Set work_Book = ThisWorkbook
    Set work_Sheet = work_Book.Worksheets("Sheet1")
    Set calculate_Sheet = work_Book.Worksheets("Calculate_Sheet")
    Set final_Sheet = work_Book.Worksheets("Final_Sheet")
    
    For Counter = 3 To Final_Row_Client_ID
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
            
        With Sheets("Calculate_Sheet").Cells(Counter, 4)
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
        final_Sheet.Cells(Counter - 1, "A").Value = calculate_Sheet.Cells(Counter, 4).Value
    Next Counter
End Sub
Private Sub BO_Name() 'Check for BO Name
    Dim Final_Row
    Dim work_Book As Workbook
    Dim work_Sheet As Worksheet
    Dim calculate_Sheet As Worksheet
    Dim final_Sheet As Worksheet
    Set work_Book = ThisWorkbook
    Set work_Sheet = work_Book.Worksheets("Sheet1")
    Set calculate_Sheet = work_Book.Worksheets("Calculate_Sheet")
    Set final_Sheet = work_Book.Worksheets("Final_Sheet")
    Final_Row = Cells(Rows.Count, "H").End(xlUp).Row
    
    For Counter = 3 To Final_Row
        
        calculate_Sheet.Cells(Counter, 1).Value = work_Sheet.Cells(Counter, "H").Value
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
        Worksheets("Final_Sheet").Cells(Counter - 1, "H").Value = calculate_Sheet.Cells(Counter, 4).Value
    Next Counter
End Sub
Private Sub URN_Type()
    Dim Final_Row_URN_Type
    Dim work_Book As Workbook
    Dim work_Sheet As Worksheet
    Dim calculate_Sheet As Worksheet
    Dim final_Sheet As Worksheet
    Set work_Book = ThisWorkbook
    Set work_Sheet = work_Book.Worksheets("Sheet1")
    Set calculate_Sheet = work_Book.Worksheets("Calculate_Sheet")
    Set final_Sheet = work_Book.Worksheets("Final_Sheet")
    Final_Row_URN_Type = Cells(Rows.Count, "J").End(xlUp).Row
    
    For Counter = 3 To Final_Row_URN_Type
    
        calculate_Sheet.Cells(Counter, 1).Value = work_Sheet.Cells(Counter, "J").Value
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
        Worksheets("Final_Sheet").Cells(Counter - 1, "J").Value = calculate_Sheet.Cells(Counter, 4).Value
    Next Counter
End Sub
Private Sub URN_Number()
    Dim Final_Row_URN_Number
    Dim work_Book As Workbook
    Dim work_Sheet As Worksheet
    Dim calculate_Sheet As Worksheet
    Dim final_Sheet As Worksheet
    Set work_Book = ThisWorkbook
    Set work_Sheet = work_Book.Worksheets("Sheet1")
    Set calculate_Sheet = work_Book.Worksheets("Calculate_Sheet")
    Set final_Sheet = work_Book.Worksheets("Final_Sheet")
    Final_Row_URN_Number = Cells(Rows.Count, "K").End(xlUp).Row
    
    For Counter = 3 To Final_Row_URN_Number
        calculate_Sheet.Cells(Counter, 1).Value = work_Sheet.Cells(Counter, "K").Value
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
        Worksheets("Final_Sheet").Cells(Counter - 1, "K").Value = calculate_Sheet.Cells(Counter, 4).Value
    Next Counter
End Sub
Private Sub Address_Type()
    Dim Final_Row_Address_Type
    Dim work_Book As Workbook
    Dim work_Sheet As Worksheet
    Dim calculate_Sheet As Worksheet
    Dim final_Sheet As Worksheet
    Set work_Book = ThisWorkbook
    Set work_Sheet = work_Book.Worksheets("Sheet1")
    Set calculate_Sheet = work_Book.Worksheets("Calculate_Sheet")
    Set final_Sheet = work_Book.Worksheets("Final_Sheet")
    Final_Row_Address_Type = Cells(Rows.Count, "O").End(xlUp).Row
    
    For Counter = 3 To Final_Row_Address_Type
        calculate_Sheet.Cells(Counter, 1).Value = work_Sheet.Cells(Counter, "O").Value
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
        Worksheets("Final_Sheet").Cells(Counter - 1, "O").Value = calculate_Sheet.Cells(Counter, 4).Value
    Next Counter
End Sub
Private Sub Address_1()
    Dim Final_Row_Address_1
    Dim work_Book As Workbook
    Dim work_Sheet As Worksheet
    Dim calculate_Sheet As Worksheet
    Dim final_Sheet As Worksheet
    Dim the_Cell_Next_Door
    'the_Cell_Next_Door = final_Sheet.Cells(Counter - 1, "Q").Value
    Set work_Book = ThisWorkbook
    Set work_Sheet = work_Book.Worksheets("Sheet1")
    Set calculate_Sheet = work_Book.Worksheets("Calculate_Sheet")
    Set final_Sheet = work_Book.Worksheets("Final_Sheet")
    Final_Row_Address_1 = work_Sheet.Cells(Rows.Count, "P").End(xlUp).Row
    
    For Counter = 3 To Final_Row_Address_1
        calculate_Sheet.Cells(Counter, 1).Value = work_Sheet.Cells(Counter, "P").Value
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
            
        If Len(calculate_Sheet.Cells(Counter, 4).Value) > 70 Then
        'MsgBox ("ee")
            calculate_Sheet.Cells(Counter, 4).TextToColumns Destination:=calculate_Sheet.Cells(Counter, "E"), DataType:=xlFixedWidth, _
                FieldInfo:=Array(Array(0, 1), Array(70, 1)), TrailingMinusNumbers:=True
            'Columns("P:P").EntireColumn.AutoFit
        End If
            
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
        
        final_Sheet.Cells(Counter - 1, "P").Value = calculate_Sheet.Cells(Counter, 5).Value
        
        If IsEmpty(final_Sheet.Cells(Counter - 1, "Q")) Then
            final_Sheet.Cells(Counter - 1, "Q").Value = calculate_Sheet.Cells(Counter, 6).Value
        End If
        
        If Not IsEmpty(final_Sheet.Cells(Counter - 1, "Q")) Then
            final_Sheet.Cells(Counter - 1, "Q").Cut final_Sheet.Cells(Counter - 1, "R")
            final_Sheet.Cells(Counter - 1, "Q").Value = calculate_Sheet.Cells(Counter, 6).Value
        End If
        
    Next Counter
End Sub
Private Sub Street_Name()
    Dim Final_Row_Street_Name
    Dim work_Book As Workbook
    Dim work_Sheet As Worksheet
    Dim calculate_Sheet As Worksheet
    Dim final_Sheet As Worksheet
    Set work_Book = ThisWorkbook
    Set work_Sheet = work_Book.Worksheets("Sheet1")
    Set calculate_Sheet = work_Book.Worksheets("Calculate_Sheet")
    Set final_Sheet = work_Book.Worksheets("Final_Sheet")
    Final_Row_Street_Name = Cells(Rows.Count, "U").End(xlUp).Row
    
    For Counter = 3 To Final_Row_Street_Name
        calculate_Sheet.Cells(Counter, 1).Value = work_Sheet.Cells(Counter, "U").Value
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
        Worksheets("Final_Sheet").Cells(Counter - 1, "U").Value = calculate_Sheet.Cells(Counter, 4).Value
    Next Counter
End Sub
Private Sub Building_No()
    Dim Final_Row_Building_No
    Dim work_Book As Workbook
    Dim work_Sheet As Worksheet
    Dim calculate_Sheet As Worksheet
    Dim final_Sheet As Worksheet
    Set work_Book = ThisWorkbook
    Set work_Sheet = work_Book.Worksheets("Sheet1")
    Set calculate_Sheet = work_Book.Worksheets("Calculate_Sheet")
    Set final_Sheet = work_Book.Worksheets("Final_Sheet")
    Final_Row_Building_No = Cells(Rows.Count, "V").End(xlUp).Row
    
    For Counter = 3 To Final_Row_Building_No
        calculate_Sheet.Cells(Counter, 1).Value = work_Sheet.Cells(Counter, "V").Value
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
        calculate_Sheet.Cells(Counter, 4).Value = "=LEFT(RC[-1], 16)"
                
        calculate_Sheet.Cells(Counter, 4).Copy
        calculate_Sheet.Cells(Counter, 5).PasteSpecial xlPasteValues
            
        With calculate_Sheet.Cells(Counter, 5)
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
        Worksheets("Final_Sheet").Cells(Counter - 1, "V").Value = calculate_Sheet.Cells(Counter, 5).Value
    Next Counter
End Sub
Private Sub PO_Box()
    Dim Final_Row_PO_Box
    Dim work_Book As Workbook
    Dim work_Sheet As Worksheet
    Dim calculate_Sheet As Worksheet
    Dim final_Sheet As Worksheet
    Set work_Book = ThisWorkbook
    Set work_Sheet = work_Book.Worksheets("Sheet1")
    Set calculate_Sheet = work_Book.Worksheets("Calculate_Sheet")
    Set final_Sheet = work_Book.Worksheets("Final_Sheet")
    Final_Row_PO_Box = Cells(Rows.Count, "W").End(xlUp).Row
    
    For Counter = 3 To Final_Row_PO_Box
        calculate_Sheet.Cells(Counter, 1).Value = work_Sheet.Cells(Counter, "W").Value
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
        calculate_Sheet.Cells(Counter, 4).Value = "=LEFT(RC[-1], 16)"
                
        calculate_Sheet.Cells(Counter, 4).Copy
        calculate_Sheet.Cells(Counter, 5).PasteSpecial xlPasteValues
            
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
        Worksheets("Final_Sheet").Cells(Counter - 1, "W").Value = calculate_Sheet.Cells(Counter, 5).Value
    Next Counter
End Sub
Private Sub Postal_Code()
    Dim Final_Row_Postal_Code
    Dim work_Book As Workbook
    Dim work_Sheet As Worksheet
    Dim calculate_Sheet As Worksheet
    Dim final_Sheet As Worksheet
    Set work_Book = ThisWorkbook
    Set work_Sheet = work_Book.Worksheets("Sheet1")
    Set calculate_Sheet = work_Book.Worksheets("Calculate_Sheet")
    Set final_Sheet = work_Book.Worksheets("Final_Sheet")
    Final_Row_Postal_Code = Cells(Rows.Count, "X").End(xlUp).Row
    
    For Counter = 3 To Final_Row_Postal_Code
        calculate_Sheet.Cells(Counter, 1).Value = work_Sheet.Cells(Counter, "X").Value
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
        calculate_Sheet.Cells(Counter, 4).Value = "=LEFT(RC[-1], 16)"
                
        calculate_Sheet.Cells(Counter, 4).Copy
        calculate_Sheet.Cells(Counter, 5).PasteSpecial xlPasteValues
            
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
        Worksheets("Final_Sheet").Cells(Counter - 1, "X").Value = calculate_Sheet.Cells(Counter, 5).Value
    Next Counter
End Sub
Private Sub Town()
    Dim Final_Row_Town
    Dim work_Book As Workbook
    Dim work_Sheet As Worksheet
    Dim calculate_Sheet As Worksheet
    Dim final_Sheet As Worksheet
    Set work_Book = ThisWorkbook
    Set work_Sheet = work_Book.Worksheets("Sheet1")
    Set calculate_Sheet = work_Book.Worksheets("Calculate_Sheet")
    Set final_Sheet = work_Book.Worksheets("Final_Sheet")
    
    Final_Row_Town = Cells(Rows.Count, "Y").End(xlUp).Row
    For Counter = 3 To Final_Row_Town
        calculate_Sheet.Cells(Counter, 1).Value = work_Sheet.Cells(Counter, "Y").Value
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
        calculate_Sheet.Cells(Counter, 4).Value = "=LEFT(RC[-1], 35)"
                
        calculate_Sheet.Cells(Counter, 4).Copy
        calculate_Sheet.Cells(Counter, 5).PasteSpecial xlPasteValues
            
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
        Worksheets("Final_Sheet").Cells(Counter - 1, "Y").Value = calculate_Sheet.Cells(Counter, 5).Value
    Next Counter
End Sub
Private Sub Province()
    Dim Final_Row_Province
    Dim work_Book As Workbook
    Dim work_Sheet As Worksheet
    Dim calculate_Sheet As Worksheet
    Dim final_Sheet As Worksheet
    Set work_Book = ThisWorkbook
    Set work_Sheet = work_Book.Worksheets("Sheet1")
    Set calculate_Sheet = work_Book.Worksheets("Calculate_Sheet")
    Set final_Sheet = work_Book.Worksheets("Final_Sheet")
    Final_Row_Province = Cells(Rows.Count, "Z").End(xlUp).Row
    
    For Counter = 3 To Final_Row_Province
        calculate_Sheet.Cells(Counter, 1).Value = work_Sheet.Cells(Counter, "Z").Value
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
        Worksheets("Final_Sheet").Cells(Counter - 1, "Z").Value = calculate_Sheet.Cells(Counter, 4).Value
    Next Counter
End Sub
Private Sub Country_Code()
    Dim Final_Row_Country_Code
    Dim work_Book As Workbook
    Dim work_Sheet As Worksheet
    Dim calculate_Sheet As Worksheet
    Dim final_Sheet As Worksheet
    Set work_Book = ThisWorkbook
    Set work_Sheet = work_Book.Worksheets("Sheet1")
    Set calculate_Sheet = work_Book.Worksheets("Calculate_Sheet")
    Set final_Sheet = work_Book.Worksheets("Final_Sheet")
    Final_Row_Country_Code = Cells(Rows.Count, "AA").End(xlUp).Row
    
    For Counter = 3 To Final_Row_Country_Code
        calculate_Sheet.Cells(Counter, 1).Value = work_Sheet.Cells(Counter, "AA").Value
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
        Worksheets("Final_Sheet").Cells(Counter - 1, "AA").Value = calculate_Sheet.Cells(Counter, 4).Value
    Next Counter
End Sub
Private Sub Account_Length() 'Check if Account Number Length is correct for User_Input CID
    Dim Final_Row_Account_Num As Long
    Dim work_Book As Workbook
    Dim work_Sheet As Worksheet
    Dim calculate_Sheet As Worksheet
    Dim final_Sheet As Worksheet
    Set work_Book = ThisWorkbook
    Set work_Sheet = work_Book.Worksheets("Sheet1")
    Set calculate_Sheet = work_Book.Worksheets("Calculate_Sheet")
    Set final_Sheet = work_Book.Worksheets("Final_Sheet")
    Dim Account_Num_Length As String
    Dim CID As String
    Dim Incorrect_Account_Length As Boolean
    CID = InputBox("Enter CID")
    Final_Row_Account_Num = work_Sheet.Cells(Rows.Count, 2).End(xlUp).Row
    
    For Counter = 3 To Final_Row_Account_Num
        If Len(work_Sheet.Cells(Counter, 2).Value) <> 5 And CID = "55P" Then
            Incorrect_Account_Length = True
            work_Sheet.Cells(Counter, 2).Interior.ColorIndex = 35
        End If
        If Len(work_Sheet.Cells(Counter, 2).Value) <> 11 And CID = "5DU" Then
            Incorrect_Account_Length = True
            work_Sheet.Cells(Counter, 2).Interior.ColorIndex = 35
        End If
    Next Counter
    If Incorrect_Account_Length Then
        MsgBox ("Cells with Incorrect Account Length have been hightlighted and sorted for you")
        Incorrect_Account_Length = False
        work_Sheet.Range("B:B").AutoFilter Field:=1, Criteria1:=RGB(204, 255 _
        , 204), Operator:=xlFilterCellColor
    End If
    
End Sub
Private Sub Create_Calculation_Sheet() 'Create a Calculate_Sheet and Final_Sheet if none exist
Attribute Create_Calculation_Sheet.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim work_Book As Workbook
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
        Worksheets(2).Name = "Calculate_Sheet"
        'ActiveSheet.Name = "Calculate_Sheet"
        'work_Sheet.Select
    End If
    
    If sheet_Final_Exists = False Then
        work_Book.Worksheets.Add After:=Worksheets("Calculate_Sheet")
        Worksheets(3).Name = "Final_Sheet"
        'ActiveSheet.Name = "Final_Sheet"
        'work_Sheet.Select
    End If
End Sub
