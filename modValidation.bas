Attribute VB_Name = "modValidation"
Option Explicit

'// these are the column constraints on the database
Public Const MAXPRODUCTNAMELEN As Integer = 50
Public Const MAXPACKAGELEN As Integer = 30
'// unit price check
Public Const MINUNITPRICE As Double = 0


Function SheetDataValidation(Target As Range) As Boolean
 
    Dim rngIntersect    As Range
    Dim intMaxChars     As Integer
    
    intMaxChars = shtControl.Range("rMaxProductNameLength")
    
    '// default the function to true
    SheetDataValidation = True
    
    Set rngIntersect = shtData.Range("rProductName")
    Set rngIntersect = rngIntersect.Resize(rngIntersect.CurrentRegion.Rows.Count - 1, rngIntersect.Columns.Count)
    Set rngIntersect = rngIntersect.Offset(1, 0) '// offset to exclude headers
    
    '// The variable rngIntersect contains the cells that will
    '// cause an alert when they are changed.
    If Not Application.Intersect(rngIntersect, Range(Target.Address)) Is Nothing Then
        If Len(Target.Text) > shtControl.Range("rMaxProductNameLength") Then
            Target.Interior.ColorIndex = 3
            MsgBox "Maximum validation length set on Control sheet for Product Name is: " & intMaxChars & _
                vbNewLine & vbNewLine & "Please enter text length less than " & intMaxChars & " characters.", vbInformation, "Validation Rule"
            Target.Activate
            SheetDataValidation = False
            Exit Function
        End If
        
        '// default the cell colour
        Target.Interior.ColorIndex = -4142
    End If
    
    Set rngIntersect = shtData.Range("rUnitPrice")
    Set rngIntersect = rngIntersect.Resize(rngIntersect.CurrentRegion.Rows.Count - 1, rngIntersect.Columns.Count)
    Set rngIntersect = rngIntersect.Offset(1, 0) '// offset to exclude headers
    
    If Not Application.Intersect(rngIntersect, Range(Target.Address)) Is Nothing Then
        If Not IsNumeric(Target.Value) Or Target.Value < MINUNITPRICE Then
            Target.Interior.ColorIndex = 3
            MsgBox "Unit price must be of a numerical value greater than 0", vbInformation, "Validation Rule"
            Target.Activate
            SheetDataValidation = False
            Exit Function
        End If
        
        '// default the cell colour
        Target.Interior.ColorIndex = -4142
    End If

    Set rngIntersect = shtData.Range("rPackage")
    Set rngIntersect = rngIntersect.Resize(rngIntersect.CurrentRegion.Rows.Count - 1, rngIntersect.Columns.Count)
    Set rngIntersect = rngIntersect.Offset(1, 0) '// offset to exclude headers
    
    If Not Application.Intersect(rngIntersect, Range(Target.Address)) Is Nothing Then
        If Len(Target.Text) > MAXPACKAGELEN Then
            Target.Interior.ColorIndex = 3
            MsgBox "Maximum validation length set on by the database is: " & MAXPACKAGELEN & _
                vbNewLine & vbNewLine & "Please enter text length less than " & MAXPACKAGELEN & " characters.", vbInformation, "Validation Rule"
            Target.Activate
            SheetDataValidation = False
            Exit Function
        End If
        
        '// default the cell colour
        Target.Interior.ColorIndex = -4142
    End If

End Function
