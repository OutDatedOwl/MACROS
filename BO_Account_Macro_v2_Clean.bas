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
    
    'Call check_Columns_For_Text
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
Private Sub Entry_Type()
    Dim work_Book As Workbook
    Dim work_Sheet As Worksheet
    Dim final_Sheet As Worksheet
    Set work_Book = ThisWorkbook
    Set work_Sheet = work_Book.Worksheets("Sheet1")
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
    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "D"), final_Sheet.Cells(2, "D")).ClearContents
    
End Sub
Private Sub Market_ID() 'Fill in Markets Column
    Dim Final_Row_Market_ID
    Dim work_Book As Workbook
    Dim work_Sheet As Worksheet
    Dim final_Sheet As Worksheet
    Set work_Book = ThisWorkbook
    Set work_Sheet = work_Book.Worksheets("Sheet1")
    Set final_Sheet = work_Book.Worksheets("Final_Sheet")
    Final_Row_Market_ID = work_Sheet.Cells(Rows.Count, 1).End(xlUp).Row
    
    final_Sheet.Select
    final_Sheet.Range(final_Sheet.Cells(Final_Row_Market_ID, "D"), final_Sheet.Cells(2, "D")).Value = "***"
    work_Sheet.Select
    
End Sub
Private Sub Entity_Type_Code()
    Dim work_Book As Workbook
    Dim work_Sheet As Worksheet
    Dim final_Sheet As Worksheet
    Set work_Book = ThisWorkbook
    Set work_Sheet = work_Book.Worksheets("Sheet1")
    Set final_Sheet = work_Book.Worksheets("Final_Sheet")
    final_Row_Data = work_Sheet.Cells(Rows.Count, 1).End(xlUp).Row
    
    work_Sheet.Range(work_Sheet.Cells(final_Row_Data, "E"), work_Sheet.Cells(2, "E")).Copy
    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "E"), final_Sheet.Cells(2, "E")).PasteSpecial xlPasteValues
    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "E"), final_Sheet.Cells(2, "E")).PasteSpecial xlPasteFormats
        
    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "F"), final_Sheet.Cells(2, "F")).Value = _
        "=IF(OR(RIGHT(TRIM(RC[-1]),1)="")"",RIGHT(TRIM(RC[-1]),1)=""."",RIGHT(TRIM(RC[-1]),1)="",""),LEFT(TRIM(RC[-1]),LEN(TRIM(RC[-1]))-1),TRIM(RC[-1]))"
    
    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "F"), final_Sheet.Cells(2, "F")).Copy
    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "E"), final_Sheet.Cells(2, "E")).PasteSpecial xlPasteValues
    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "F"), final_Sheet.Cells(2, "F")).ClearContents
    
End Sub
Private Sub Check_CID() 'Checking for CID
    Dim work_Book As Workbook
    Dim work_Sheet As Worksheet
    Dim final_Sheet As Worksheet
    Set work_Book = ThisWorkbook
    Set work_Sheet = work_Book.Worksheets("Sheet1")
    Set final_Sheet = work_Book.Worksheets("Final_Sheet")
    final_Row_Data = work_Sheet.Cells(Rows.Count, 1).End(xlUp).Row
    
    work_Sheet.Range(work_Sheet.Cells(final_Row_Data, "A"), work_Sheet.Cells(3, "A")).Copy
    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "A"), final_Sheet.Cells(2, "A")).PasteSpecial xlPasteValues
    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "A"), final_Sheet.Cells(2, "A")).PasteSpecial xlPasteFormats
        
    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "B"), final_Sheet.Cells(2, "B")).Value = _
        "=IF(OR(RIGHT(TRIM(RC[-1]),1)="")"",RIGHT(TRIM(RC[-1]),1)=""."",RIGHT(TRIM(RC[-1]),1)="",""),LEFT(TRIM(RC[-1]),LEN(TRIM(RC[-1]))-1),TRIM(RC[-1]))"
    
    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "B"), final_Sheet.Cells(2, "B")).Copy
    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "A"), final_Sheet.Cells(2, "A")).PasteSpecial xlPasteValues
    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "B"), final_Sheet.Cells(2, "B")).ClearContents
    
End Sub
Private Sub BO_Name() 'Check for BO Name
    Dim work_Book As Workbook
    Dim work_Sheet As Worksheet
    Dim final_Sheet As Worksheet
    Set work_Book = ThisWorkbook
    Set work_Sheet = work_Book.Worksheets("Sheet1")
    Set final_Sheet = work_Book.Worksheets("Final_Sheet")
    final_Row_Data = work_Sheet.Cells(Rows.Count, "H").End(xlUp).Row
    
    work_Sheet.Range(work_Sheet.Cells(final_Row_Data, "H"), work_Sheet.Cells(2, "H")).Copy
    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "H"), final_Sheet.Cells(2, "H")).PasteSpecial xlPasteValues
    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "H"), final_Sheet.Cells(2, "H")).PasteSpecial xlPasteFormats
        
    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "I"), final_Sheet.Cells(2, "I")).Value = _
        "=IF(OR(RIGHT(TRIM(RC[-1]),1)="")"",RIGHT(TRIM(RC[-1]),1)=""."",RIGHT(TRIM(RC[-1]),1)="",""),LEFT(TRIM(RC[-1]),LEN(TRIM(RC[-1]))-1),TRIM(RC[-1]))"
    
    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "I"), final_Sheet.Cells(2, "I")).Copy
    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "H"), final_Sheet.Cells(2, "H")).PasteSpecial xlPasteValues
    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "I"), final_Sheet.Cells(2, "I")).ClearContents
    
End Sub
Private Sub URN_Type()
    Dim work_Book As Workbook
    Dim work_Sheet As Worksheet
    Dim final_Sheet As Worksheet
    Set work_Book = ThisWorkbook
    Set work_Sheet = work_Book.Worksheets("Sheet1")
    Set final_Sheet = work_Book.Worksheets("Final_Sheet")
    final_Row_Data = work_Sheet.Cells(Rows.Count, "J").End(xlUp).Row
    
    If final_Row_Data <> 1 Then
        work_Sheet.Range(work_Sheet.Cells(final_Row_Data, "J"), work_Sheet.Cells(2, "J")).Copy
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "J"), final_Sheet.Cells(2, "J")).PasteSpecial xlPasteValues
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "J"), final_Sheet.Cells(2, "J")).PasteSpecial xlPasteFormats
        
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "J"), final_Sheet.Cells(2, "J")).Replace What:="N/A", Replacement:="", LookAt:= _
            xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
            
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "K"), final_Sheet.Cells(2, "K")).Value = _
            "=IF(OR(RIGHT(TRIM(RC[-1]),1)="")"",RIGHT(TRIM(RC[-1]),1)=""."",RIGHT(TRIM(RC[-1]),1)="",""),LEFT(TRIM(RC[-1]),LEN(TRIM(RC[-1]))-1),TRIM(RC[-1]))"
        
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "K"), final_Sheet.Cells(2, "K")).Copy
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "J"), final_Sheet.Cells(2, "J")).PasteSpecial xlPasteValues
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "K"), final_Sheet.Cells(2, "K")).ClearContents
        
    End If
End Sub
Private Sub URN_Number()
    Dim work_Book As Workbook
    Dim work_Sheet As Worksheet
    Dim final_Sheet As Worksheet
    Set work_Book = ThisWorkbook
    Set work_Sheet = work_Book.Worksheets("Sheet1")
    Set final_Sheet = work_Book.Worksheets("Final_Sheet")
    final_Row_Data = work_Sheet.Cells(Rows.Count, "K").End(xlUp).Row
    
    If final_Row_Data <> 1 Then
        work_Sheet.Range(work_Sheet.Cells(final_Row_Data, "K"), work_Sheet.Cells(2, "K")).Copy
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "K"), final_Sheet.Cells(2, "K")).PasteSpecial xlPasteValues
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "K"), final_Sheet.Cells(2, "K")).PasteSpecial xlPasteFormats
'
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "K"), final_Sheet.Cells(2, "K")).Replace What:="N/A", Replacement:="", LookAt:= _
            xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
            
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "L"), final_Sheet.Cells(2, "L")).Value = _
            "=IF(OR(RIGHT(TRIM(RC[-1]),1)="")"",RIGHT(TRIM(RC[-1]),1)=""."",RIGHT(TRIM(RC[-1]),1)="",""),LEFT(TRIM(RC[-1]),LEN(TRIM(RC[-1]))-1),TRIM(RC[-1]))"
        
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "L"), final_Sheet.Cells(2, "L")).Copy
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "K"), final_Sheet.Cells(2, "K")).PasteSpecial xlPasteValues
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "L"), final_Sheet.Cells(2, "L")).ClearContents
        
    End If
End Sub
Private Sub Address_Type()
    Dim work_Book As Workbook
    Dim work_Sheet As Worksheet
    Dim final_Sheet As Worksheet
    Set work_Book = ThisWorkbook
    Set work_Sheet = work_Book.Worksheets("Sheet1")
    Set final_Sheet = work_Book.Worksheets("Final_Sheet")
    final_Row_Data = work_Sheet.Cells(Rows.Count, "O").End(xlUp).Row
    
    If final_Row_Data <> 1 Then
        work_Sheet.Range(work_Sheet.Cells(final_Row_Data, "O"), work_Sheet.Cells(2, "O")).Copy
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "O"), final_Sheet.Cells(2, "O")).PasteSpecial xlPasteValues
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "O"), final_Sheet.Cells(2, "O")).PasteSpecial xlPasteFormats
            
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "P"), final_Sheet.Cells(2, "P")).Value = _
            "=IF(OR(RIGHT(TRIM(RC[-1]),1)="")"",RIGHT(TRIM(RC[-1]),1)=""."",RIGHT(TRIM(RC[-1]),1)="",""),LEFT(TRIM(RC[-1]),LEN(TRIM(RC[-1]))-1),TRIM(RC[-1]))"
        
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "P"), final_Sheet.Cells(2, "P")).Copy
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "O"), final_Sheet.Cells(2, "O")).PasteSpecial xlPasteValues
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "P"), final_Sheet.Cells(2, "P")).ClearContents
        
    End If
End Sub
Private Sub Address()
    Dim work_Book As Workbook
    Dim work_Sheet As Worksheet
    Dim final_Sheet As Worksheet
    Set work_Book = ThisWorkbook
    Set work_Sheet = work_Book.Worksheets("Sheet1")
    Set final_Sheet = work_Book.Worksheets("Final_Sheet")
    Dim account_Address_1_Range As Range
    Dim account_Address_2_Range As Range
    Dim account_Address_3_Range As Range
    Dim account_Address_4_Range As Range
    Dim account_Address_5_Range As Range
    Dim column_A_To_AA As Range
    Dim account_Address_Cell_1 As Range
    Dim account_Address_Cell_2 As Range
    Dim account_Address_Cell_3 As Range
    Dim account_Address_Cell_4 As Range
    Dim account_Address_Cell_5 As Range
    final_Row_Data = work_Sheet.Cells(Rows.Count, "P").End(xlUp).Row 'final_Row for Address 1
    Set account_Address_1_Range = work_Sheet.Range(work_Sheet.Cells(2, "P"), work_Sheet.Cells(final_Row_Data, "P")) 'Set up Address Range 1 on Sheet1
    
    If final_Row_Data <> 1 Then 'Address 1 final_Row_Data cannot be 1
        work_Sheet.Range(work_Sheet.Cells(final_Row_Data, "P"), work_Sheet.Cells(2, "P")).Copy
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "P"), final_Sheet.Cells(2, "P")).PasteSpecial xlPasteValues
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "P"), final_Sheet.Cells(2, "P")).PasteSpecial xlPasteFormats
            
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
        
 '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
    final_Row_Data = work_Sheet.Cells(Rows.Count, "Q").End(xlUp).Row 'final_Row for Address 2
    
    If final_Row_Data <> 1 Then 'Address 2, final_Row_Data cannot be 1
        Set account_Address_2_Range = final_Sheet.Range(final_Sheet.Cells(2, "Q"), final_Sheet.Cells(final_Row_Data, "Q")) 'Set up Address Range 2
    
        For Each account_Address_Cell_2 In account_Address_2_Range.Cells
            If Not IsEmpty(account_Address_Cell_2.Value) Then 'If not empty then load into adjacent column ->
                work_Sheet.Range(work_Sheet.Cells(account_Address_Cell_2.Row, "Q"), work_Sheet.Cells(account_Address_Cell_2.Row, "Q")).Copy
    
                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_2.Row, "R"), final_Sheet.Cells(account_Address_Cell_2.Row, "R")).PasteSpecial xlPasteValues
                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_2.Row, "R"), final_Sheet.Cells(account_Address_Cell_2.Row, "R")).PasteSpecial xlPasteFormats
    
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
    
                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_2.Row, "R"), final_Sheet.Cells(account_Address_Cell_2.Row, "R")).Value = _
                    "=IF(OR(RIGHT(TRIM(RC[-1]),1)="")"",RIGHT(TRIM(RC[-1]),1)=""."",RIGHT(TRIM(RC[-1]),1)="",""),LEFT(TRIM(RC[-1]),LEN(TRIM(RC[-1]))-1),TRIM(RC[-1]))"
    
                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_2.Row, "R"), final_Sheet.Cells(account_Address_Cell_2.Row, "R")).Copy
                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_2.Row, "Q"), final_Sheet.Cells(account_Address_Cell_2.Row, "Q")).PasteSpecial xlPasteValues
                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_2.Row, "R"), final_Sheet.Cells(account_Address_Cell_2.Row, "R")).ClearContents
            End If
        Next account_Address_Cell_2
        
        For Each account_Address_Cell_2 In account_Address_2_Range.Cells 'Address 2 Text to Columns
            If Len(account_Address_Cell_2.Value) > 70 Then
                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_2.Row, "Q"), final_Sheet.Cells(account_Address_Cell_2.Row, "Q")).TextToColumns Destination:= _
                    final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_2.Row, "Q"), final_Sheet.Cells(account_Address_Cell_2.Row, "Q")), DataType:=xlFixedWidth, _
                        FieldInfo:=Array(Array(0, 1), Array(70, 1)), TrailingMinusNumbers:=True
            End If
        Next account_Address_Cell_2
    End If
    
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
    final_Row_Data = work_Sheet.Cells(Rows.Count, "R").End(xlUp).Row 'final_Row for Address 3
    
    If final_Row_Data <> 1 Then 'Address 3,final_Row_Data cannot be 1
        Set account_Address_3_Range = final_Sheet.Range(final_Sheet.Cells(2, "R"), final_Sheet.Cells(final_Row_Data, "R")) 'Set up Address Range 3
        
        For Each account_Address_Cell_3 In account_Address_3_Range.Cells
            If Not IsEmpty(account_Address_Cell_3.Value) Then 'If not empty then load into adjacent column ->
                work_Sheet.Range(work_Sheet.Cells(account_Address_Cell_3.Row, "R"), work_Sheet.Cells(account_Address_Cell_3.Row, "R")).Copy

                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_3.Row, "S"), final_Sheet.Cells(account_Address_Cell_3.Row, "S")).PasteSpecial xlPasteValues
                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_3.Row, "S"), final_Sheet.Cells(account_Address_Cell_3.Row, "S")).PasteSpecial xlPasteFormats

                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_3.Row, "T"), final_Sheet.Cells(account_Address_Cell_3.Row, "T")).Value = _
                    "=IF(OR(RIGHT(TRIM(RC[-1]),1)="")"",RIGHT(TRIM(RC[-1]),1)=""."",RIGHT(TRIM(RC[-1]),1)="",""),LEFT(TRIM(RC[-1]),LEN(TRIM(RC[-1]))-1),TRIM(RC[-1]))"

                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_3.Row, "T"), final_Sheet.Cells(account_Address_Cell_3.Row, "T")).Copy
                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_3.Row, "S"), final_Sheet.Cells(account_Address_Cell_3.Row, "S")).PasteSpecial xlPasteValues
                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_3.Row, "T"), final_Sheet.Cells(account_Address_Cell_3.Row, "T")).ClearContents
            End If
            If IsEmpty(account_Address_Cell_3.Value) Then  'Checking if cell contains text from Address 2, if empty then load into correct column
                work_Sheet.Range(work_Sheet.Cells(account_Address_Cell_3.Row, "R"), work_Sheet.Cells(account_Address_Cell_3.Row, "R")).Copy

                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_3.Row, "R"), final_Sheet.Cells(account_Address_Cell_3.Row, "R")).PasteSpecial xlPasteValues
                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_3.Row, "R"), final_Sheet.Cells(account_Address_Cell_3.Row, "R")).PasteSpecial xlPasteFormats

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
    
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    If final_Row_Data <> 1 Then 'Address 4,final_Row_Data cannot be 1
        Set account_Address_4_Range = final_Sheet.Range(final_Sheet.Cells(2, "S"), final_Sheet.Cells(final_Row_Data, "S")) 'Set up Address Range 4
        
        For Each account_Address_Cell_4 In account_Address_4_Range.Cells
            If Not IsEmpty(account_Address_Cell_4.Value) Then 'If not empty then load into adjacent column ->
                work_Sheet.Range(work_Sheet.Cells(account_Address_Cell_4.Row, "S"), work_Sheet.Cells(account_Address_Cell_4.Row, "S")).Copy

                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_4.Row, "S"), final_Sheet.Cells(account_Address_Cell_4.Row, "S")).PasteSpecial xlPasteValues
                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_4.Row, "S"), final_Sheet.Cells(account_Address_Cell_4.Row, "S")).PasteSpecial xlPasteFormats

                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_4.Row, "T"), final_Sheet.Cells(account_Address_Cell_4.Row, "T")).Value = _
                    "=IF(OR(RIGHT(TRIM(RC[-1]),1)="")"",RIGHT(TRIM(RC[-1]),1)=""."",RIGHT(TRIM(RC[-1]),1)="",""),LEFT(TRIM(RC[-1]),LEN(TRIM(RC[-1]))-1),TRIM(RC[-1]))"

                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_4.Row, "T"), final_Sheet.Cells(account_Address_Cell_4.Row, "T")).Copy
                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_4.Row, "S"), final_Sheet.Cells(account_Address_Cell_4.Row, "S")).PasteSpecial xlPasteValues
                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_4.Row, "T"), final_Sheet.Cells(account_Address_Cell_4.Row, "T")).ClearContents
            End If
            If IsEmpty(account_Address_Cell_4.Value) Then  'Checking if cell contains text from Address 2, if empty then load into correct column
                work_Sheet.Range(work_Sheet.Cells(account_Address_Cell_4.Row, "R"), work_Sheet.Cells(account_Address_Cell_4.Row, "R")).Copy

                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_4.Row, "S"), final_Sheet.Cells(account_Address_Cell_4.Row, "S")).Value = _
                    "=IF(OR(RIGHT(TRIM(RC[-1]),1)="")"",RIGHT(TRIM(RC[-1]),1)=""."",RIGHT(TRIM(RC[-1]),1)="",""),LEFT(TRIM(RC[-1]),LEN(TRIM(RC[-1]))-1),TRIM(RC[-1]))"

                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_4.Row, "S"), final_Sheet.Cells(account_Address_Cell_4.Row, "S")).Copy
                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_4.Row, "R"), final_Sheet.Cells(account_Address_Cell_4.Row, "R")).PasteSpecial xlPasteValues
                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_4.Row, "S"), final_Sheet.Cells(account_Address_Cell_4.Row, "S")).ClearContents
            End If
        Next account_Address_Cell_4
        
        For Each account_Address_Cell_4 In account_Address_4_Range.Cells 'Address 4 Text to Columns
            If Len(account_Address_Cell_4.Value) > 70 Then
                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_4.Row, "S"), final_Sheet.Cells(account_Address_Cell_4.Row, "S")).TextToColumns Destination:= _
                    final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_4.Row, "S"), final_Sheet.Cells(account_Address_Cell_4.Row, "S")), DataType:=xlFixedWidth, _
                        FieldInfo:=Array(Array(0, 1), Array(70, 1)), TrailingMinusNumbers:=True
            End If
        Next account_Address_Cell_4
    End If
    
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    If final_Row_Data <> 1 Then 'Address 5,final_Row_Data cannot be 1
        Set account_Address_5_Range = final_Sheet.Range(final_Sheet.Cells(2, "T"), final_Sheet.Cells(final_Row_Data, "T")) 'Set up Address Range 5
        
        For Each account_Address_Cell_5 In account_Address_5_Range.Cells
            If Not IsEmpty(account_Address_Cell_5.Value) Then 'If not empty then load into adjacent column ->
                If Len(account_Address_Cell_5) > 70 Then
                    account_Address_Cell_5 = Left(account_Address_Cell_5, 70)
                End If

            End If
            If IsEmpty(account_Address_Cell_5.Value) Then  'Checking if cell contains text from Address 2, if empty then load into correct column
                work_Sheet.Range(work_Sheet.Cells(account_Address_Cell_5.Row, "T"), work_Sheet.Cells(account_Address_Cell_5.Row, "T")).Copy

                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_5.Row, "U"), final_Sheet.Cells(account_Address_Cell_5.Row, "U")).Value = _
                    "=IF(OR(RIGHT(TRIM(RC[-1]),1)="")"",RIGHT(TRIM(RC[-1]),1)=""."",RIGHT(TRIM(RC[-1]),1)="",""),LEFT(TRIM(RC[-1]),LEN(TRIM(RC[-1]))-1),TRIM(RC[-1]))"

                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_5.Row, "U"), final_Sheet.Cells(account_Address_Cell_5.Row, "U")).Copy
                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_5.Row, "T"), final_Sheet.Cells(account_Address_Cell_5.Row, "T")).PasteSpecial xlPasteValues
                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_5.Row, "U"), final_Sheet.Cells(account_Address_Cell_5.Row, "U")).ClearContents
            End If
        Next account_Address_Cell_5
        
        For Each account_Address_Cell_5 In account_Address_5_Range.Cells 'Address 5 Text to Columns
            If Len(account_Address_Cell_5.Value) > 70 Then
                final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_5.Row, "T"), final_Sheet.Cells(account_Address_Cell_5.Row, "T")).TextToColumns Destination:= _
                    final_Sheet.Range(final_Sheet.Cells(account_Address_Cell_5.Row, "T"), final_Sheet.Cells(account_Address_Cell_5.Row, "T")), DataType:=xlFixedWidth, _
                        FieldInfo:=Array(Array(0, 1), Array(70, 1)), TrailingMinusNumbers:=True
            End If
        Next account_Address_Cell_5
    End If
    
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

    column_A_To_AA.Replace What:="^", Replacement:="", LookAt:= _
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
    Dim work_Book As Workbook
    Dim work_Sheet As Worksheet
    Dim final_Sheet As Worksheet
    Set work_Book = ThisWorkbook
    Set work_Sheet = work_Book.Worksheets("Sheet1")
    Set final_Sheet = work_Book.Worksheets("Final_Sheet")
    final_Row_Data = Cells(Rows.Count, "U").End(xlUp).Row
    
    If final_Row_Data <> 1 Then
        work_Sheet.Range(work_Sheet.Cells(final_Row_Data, "U"), work_Sheet.Cells(2, "U")).Copy
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "U"), final_Sheet.Cells(2, "U")).PasteSpecial xlPasteValues
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "U"), final_Sheet.Cells(2, "U")).PasteSpecial xlPasteFormats
            
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "V"), final_Sheet.Cells(2, "V")).Value = _
            "=IF(OR(RIGHT(TRIM(RC[-1]),1)="")"",RIGHT(TRIM(RC[-1]),1)=""."",RIGHT(TRIM(RC[-1]),1)="",""),LEFT(TRIM(RC[-1]),LEN(TRIM(RC[-1]))-1),TRIM(RC[-1]))"
        
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "V"), final_Sheet.Cells(2, "V")).Copy
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "U"), final_Sheet.Cells(2, "U")).PasteSpecial xlPasteValues
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "V"), final_Sheet.Cells(2, "V")).ClearContents
        
    End If
End Sub
Private Sub Building_No()
    Dim work_Book As Workbook
    Dim work_Sheet As Worksheet
    Dim final_Sheet As Worksheet
    Set work_Book = ThisWorkbook
    Set work_Sheet = work_Book.Worksheets("Sheet1")
    Set final_Sheet = work_Book.Worksheets("Final_Sheet")
    final_Row_Data = work_Sheet.Cells(Rows.Count, "V").End(xlUp).Row
    
    If final_Row_Data <> 1 Then
        work_Sheet.Range(work_Sheet.Cells(final_Row_Data, "V"), work_Sheet.Cells(2, "V")).Copy
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "V"), final_Sheet.Cells(2, "V")).PasteSpecial xlPasteValues
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "V"), final_Sheet.Cells(2, "V")).PasteSpecial xlPasteFormats
            
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "W"), final_Sheet.Cells(2, "W")).Value = _
            "=IF(OR(RIGHT(TRIM(RC[-1]),1)="")"",RIGHT(TRIM(RC[-1]),1)=""."",RIGHT(TRIM(RC[-1]),1)="",""),LEFT(TRIM(RC[-1]),LEN(TRIM(RC[-1]))-1),TRIM(RC[-1]))"
        
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "W"), final_Sheet.Cells(2, "W")).Copy
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "V"), final_Sheet.Cells(2, "V")).PasteSpecial xlPasteValues
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "W"), final_Sheet.Cells(2, "W")).ClearContents
        
    End If
End Sub
Private Sub PO_Box()
    Dim work_Book As Workbook
    Dim work_Sheet As Worksheet
    Dim final_Sheet As Worksheet
    Set work_Book = ThisWorkbook
    Set work_Sheet = work_Book.Worksheets("Sheet1")
    Set final_Sheet = work_Book.Worksheets("Final_Sheet")
    final_Row_Data = work_Sheet.Cells(Rows.Count, "W").End(xlUp).Row
    
    If final_Row_Data <> 1 Then
        work_Sheet.Range(work_Sheet.Cells(final_Row_Data, "W"), work_Sheet.Cells(2, "W")).Copy
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "W"), final_Sheet.Cells(2, "W")).PasteSpecial xlPasteValues
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "W"), final_Sheet.Cells(2, "W")).PasteSpecial xlPasteFormats
            
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "X"), final_Sheet.Cells(2, "X")).Value = _
            "=IF(OR(RIGHT(TRIM(RC[-1]),1)="")"",RIGHT(TRIM(RC[-1]),1)=""."",RIGHT(TRIM(RC[-1]),1)="",""),LEFT(TRIM(RC[-1]),LEN(TRIM(RC[-1]))-1),TRIM(RC[-1]))"
        
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "X"), final_Sheet.Cells(2, "X")).Copy
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "W"), final_Sheet.Cells(2, "W")).PasteSpecial xlPasteValues
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "X"), final_Sheet.Cells(2, "X")).ClearContents

    End If
End Sub
Private Sub Postal_Code()
    Dim work_Book As Workbook
    Dim work_Sheet As Worksheet
    Dim final_Sheet As Worksheet
    Set work_Book = ThisWorkbook
    Set work_Sheet = work_Book.Worksheets("Sheet1")
    Set final_Sheet = work_Book.Worksheets("Final_Sheet")
    final_Row_Data = work_Sheet.Cells(Rows.Count, "X").End(xlUp).Row
    
    If final_Row_Data <> 1 Then
        work_Sheet.Range(work_Sheet.Cells(final_Row_Data, "X"), work_Sheet.Cells(2, "X")).Copy
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "X"), final_Sheet.Cells(2, "X")).PasteSpecial xlPasteValues
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "X"), final_Sheet.Cells(2, "X")).PasteSpecial xlPasteFormats
            
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "Y"), final_Sheet.Cells(2, "Y")).Value = _
            "=IF(OR(RIGHT(TRIM(RC[-1]),1)="")"",RIGHT(TRIM(RC[-1]),1)=""."",RIGHT(TRIM(RC[-1]),1)="",""),LEFT(TRIM(RC[-1]),LEN(TRIM(RC[-1]))-1),TRIM(RC[-1]))"
        
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "Y"), final_Sheet.Cells(2, "Y")).Copy
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "X"), final_Sheet.Cells(2, "X")).PasteSpecial xlPasteValues
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "Y"), final_Sheet.Cells(2, "Y")).ClearContents
        
    End If
End Sub
Private Sub Town()
    Dim work_Book As Workbook
    Dim work_Sheet As Worksheet
    Dim final_Sheet As Worksheet
    Set work_Book = ThisWorkbook
    Set work_Sheet = work_Book.Worksheets("Sheet1")
    Set final_Sheet = work_Book.Worksheets("Final_Sheet")
    final_Row_Data = work_Sheet.Cells(Rows.Count, "Y").End(xlUp).Row
    
    
    If final_Row_Data <> 1 Then
        work_Sheet.Range(work_Sheet.Cells(final_Row_Data, "Y"), work_Sheet.Cells(2, "Y")).Copy
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "Y"), final_Sheet.Cells(2, "Y")).PasteSpecial xlPasteValues
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "Y"), final_Sheet.Cells(2, "Y")).PasteSpecial xlPasteFormats
            
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "Z"), final_Sheet.Cells(2, "Z")).Value = _
            "=IF(OR(RIGHT(TRIM(RC[-1]),1)="")"",RIGHT(TRIM(RC[-1]),1)=""."",RIGHT(TRIM(RC[-1]),1)="",""),LEFT(TRIM(RC[-1]),LEN(TRIM(RC[-1]))-1),TRIM(RC[-1]))"
        
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "Z"), final_Sheet.Cells(2, "Z")).Copy
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "Y"), final_Sheet.Cells(2, "Y")).PasteSpecial xlPasteValues
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "Z"), final_Sheet.Cells(2, "Z")).ClearContents
        
    End If
End Sub
Private Sub Province()
    Dim work_Book As Workbook
    Dim work_Sheet As Worksheet
    Dim final_Sheet As Worksheet
    Set work_Book = ThisWorkbook
    Set work_Sheet = work_Book.Worksheets("Sheet1")
    Set final_Sheet = work_Book.Worksheets("Final_Sheet")
    final_Row_Data = work_Sheet.Cells(Rows.Count, "Z").End(xlUp).Row
    
    work_Sheet.Range(work_Sheet.Cells(final_Row_Data, "Z"), work_Sheet.Cells(2, "Z")).Copy
    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "Z"), final_Sheet.Cells(2, "Z")).PasteSpecial xlPasteValues
    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "Z"), final_Sheet.Cells(2, "Z")).PasteSpecial xlPasteFormats
        
    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "AA"), final_Sheet.Cells(2, "AA")).Value = _
        "=IF(OR(RIGHT(TRIM(RC[-1]),1)="")"",RIGHT(TRIM(RC[-1]),1)=""."",RIGHT(TRIM(RC[-1]),1)="",""),LEFT(TRIM(RC[-1]),LEN(TRIM(RC[-1]))-1),TRIM(RC[-1]))"
    
    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "AA"), final_Sheet.Cells(2, "AA")).Copy
    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "Z"), final_Sheet.Cells(2, "Z")).PasteSpecial xlPasteValues
    final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "AA"), final_Sheet.Cells(2, "AA")).ClearContents
    
End Sub
Private Sub Country_Code()
    Dim work_Book As Workbook
    Dim work_Sheet As Worksheet
    Dim final_Sheet As Worksheet
    Set work_Book = ThisWorkbook
    Set work_Sheet = work_Book.Worksheets("Sheet1")
    Set final_Sheet = work_Book.Worksheets("Final_Sheet")
    final_Row_Data = work_Sheet.Cells(Rows.Count, "AA").End(xlUp).Row
    
    If final_Row_Data <> 1 Then
        work_Sheet.Range(work_Sheet.Cells(final_Row_Data, "AA"), work_Sheet.Cells(2, "AA")).Copy
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "AA"), final_Sheet.Cells(2, "AA")).PasteSpecial xlPasteValues
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "AA"), final_Sheet.Cells(2, "AA")).PasteSpecial xlPasteFormats
            
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "AB"), final_Sheet.Cells(2, "AB")).Value = _
            "=IF(OR(RIGHT(TRIM(RC[-1]),1)="")"",RIGHT(TRIM(RC[-1]),1)=""."",RIGHT(TRIM(RC[-1]),1)="",""),LEFT(TRIM(RC[-1]),LEN(TRIM(RC[-1]))-1),TRIM(RC[-1]))"
        
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "AB"), final_Sheet.Cells(2, "AB")).Copy
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "AA"), final_Sheet.Cells(2, "AA")).PasteSpecial xlPasteValues
        final_Sheet.Range(final_Sheet.Cells(final_Row_Data, "AB"), final_Sheet.Cells(2, "AB")).ClearContents
        
    End If
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
Attribute Create_Final_Sheet.VB_ProcData.VB_Invoke_Func = " \n14"
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
        work_Book.Worksheets.Add After:=Worksheets("Sheet1").Name = "Final_Sheet"
    End If
End Sub
