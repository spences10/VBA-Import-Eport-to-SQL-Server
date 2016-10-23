Attribute VB_Name = "modImportExport"
Option Explicit

'// references
'// Microsoft Active X Data Objects 2.6 Library

'// General use module level variables
Private m_cn                As ADODB.Connection
Private m_rs                As ADODB.Recordset

Private m_strSQL            As String

Private m_strServerName     As String
Private m_strDatabaseName   As String
Private m_strUserID         As String
Private m_strPassword       As String

Sub Auto_Open()
    
    On Error GoTo CatchError
    
    '// default to the control page
    shtControl.Activate
    
    '// Collect connection variables
    Call ConnectionVariables
    
    '// clear down cbo items before proceeding
    shtControl.cboCompany.Clear
    
    If Not InitializeConnection Then
        Exit Sub
    End If
    
    Set m_rs = New ADODB.Recordset
    '// populate recordset with company names
    m_rs.Open "EXEC spSuppliersList", m_cn, adOpenStatic, adLockReadOnly
    
    '// instead of looping through recordset
    '// assign recordset via GetRows then use .Transpose
    '// this will help bind more than one column to the control
    shtControl.cboCompany.List = Application.Transpose(m_rs.GetRows)
      
    '// dont default to the first item in the list
    shtControl.cboCompany.ListIndex = -1

ExitSub:
    Call TerminateConnection
    Exit Sub

CatchError:
    MsgBox "Error Naumber: " & Err.Number & vbNewLine & vbNewLine & _
        "Error Message: " & Err.Description, vbCritical, "Run Time Error"
    GoTo ExitSub

End Sub

Sub Auto_Close()

    Dim rngBoardersRange    As Range
    
    Call TerminateConnection
    
    '// clear up shtData
    Set rngBoardersRange = shtData.Range("rProductID").CurrentRegion
    Application.EnableEvents = False
    rngBoardersRange.ClearContents
    Call RangeBorders(rngBoardersRange, True)
    Application.EnableEvents = True
    
    ThisWorkbook.Save
    
End Sub


Sub ConnectionVariables()

    '// collect conncetion variables
    m_strServerName = shtControl.Range("rServerName")
    m_strDatabaseName = shtControl.Range("rDatabaseName")
    m_strUserID = shtControl.Range("rUserID")
    m_strPassword = shtControl.Range("rPassword")
    
End Sub

'// attempts to set up a database connection
'// wih the credentials supplied on shtControl
Function InitializeConnection() As Boolean
    
    On Error GoTo CatchError
    
    If m_strServerName = "" And m_strDatabaseName = "" Then
        Call ConnectionVariables
    End If
    
    '// use global variables in connecting
    Set m_cn = New ADODB.Connection
    m_cn.Open "Driver={SQL Server};Server=" & m_strServerName & ";Database=" & m_strDatabaseName & _
        ";Uid=" & m_strUserID & ";Pwd=" & m_strPassword & ";"

    If Not m_cn Is Nothing Then
        If m_cn.State = adStateClosed Then m_cn.Open
    End If
    
    InitializeConnection = True

ExitSub:
    Exit Function

CatchError:
    
    InitializeConnection = False
    
    Select Case Err.Description
        
        Case "[Microsoft][ODBC SQL Server Driver][DBNETLIB]SQL Server does not exist or access denied."
            MsgBox "You do not seem to have the correct SQL Server name please check " & _
                "the details on the Control sheet are correct before attempting again." & vbNewLine & vbNewLine & _
                "Error Naumber: " & Err.Number & vbNewLine & vbNewLine & _
                "Error Message: " & Err.Description, vbCritical, "Run Time Error"
                
        Case "[Microsoft][ODBC SQL Server Driver][SQL Server]Cannot open database " & Chr(34) & m_strDatabaseName & Chr(34) & " requested by the login. The login failed."
            MsgBox "You do not seem to have the correct database name please check " & _
                "the details on the Control sheet are correct before attempting again." & vbNewLine & vbNewLine & _
                "Error Naumber: " & Err.Number & vbNewLine & vbNewLine & _
                "Error Message: " & Err.Description, vbCritical, "Run Time Error"
                
        Case Else
            MsgBox "Error Naumber: " & Err.Number & vbNewLine & vbNewLine & _
                "Error Message: " & Err.Description, vbCritical, "Run Time Error"
                
    End Select
    
    Call TerminateConnection
    GoTo ExitSub
    
End Function

Sub TerminateConnection()
    
    Set m_rs = Nothing
    Set m_cn = Nothing

End Sub

Sub ImportData()
    
    Dim intColIndex     As Integer
    Dim rngUsedRange    As Range
    
    Dim rngMyCell       As Range
    
    On Error GoTo CatchError
    
    '// niitalize the database connection
    Call InitializeConnection
    
    '// build the query string for the recordset,
    '// if it is required that the name length be truncated from the server then uncomment the last section of the next line
    m_strSQL = "EXEC spGetAllSupplierProducts " & shtControl.cboCompany.Value '// & ", " & shtControl.Range("rMaxProductNameLength")
    '// create recordset
    Set m_rs = New ADODB.Recordset
    m_rs.Open m_strSQL, m_cn, adOpenStatic
    '// setup shtData for recordset data
    Set rngUsedRange = shtData.Range("rProductID").CurrentRegion
    Call RangeBorders(rngUsedRange, True)
    
    '// disable events so as not to trigger sheet validation
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    rngUsedRange.ClearContents
    
    '// paste recordset
    With shtData.Range("rProductID").Offset(1, 0)
        .CopyFromRecordset m_rs
    End With
    
    '// add in headers
    For intColIndex = 0 To m_rs.Fields.Count - 1
        shtData.Range("rProductID").Offset(0, intColIndex).Value = m_rs.Fields(intColIndex).Name
        shtData.Range("rProductID").Offset(0, intColIndex).Font.Bold = True
        shtData.Columns(intColIndex + 1).AutoFit
    Next
    
    '// enable events again
    Application.EnableEvents = True
    shtData.Activate
    
    '// validate it agains rules
    For Each rngMyCell In rngUsedRange
        If Not SheetDataValidation(rngMyCell) Then
            '// mark the cell in the validation module
        End If
    Next rngMyCell
    
    Application.ScreenUpdating = True
        
    '// hide the ID column and the repeated company name
    shtData.Columns("B:C").Hidden = True
    '// pretty it up with boarders
    Set rngUsedRange = shtData.Range("rProductID").CurrentRegion
    Call RangeBorders(rngUsedRange, False)
    '// display sheet to user


ExitSub:
    Call TerminateConnection
    Exit Sub

CatchError:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "Error Naumber: " & Err.Number & vbNewLine & vbNewLine & _
        "Error Message: " & Err.Description, vbCritical, "Run Time Error"
    GoTo ExitSub
    
End Sub


Sub ExportData()
    
    Dim rngValidationRange  As Range
    Dim rngMyCell           As Range
    Dim rngImport           As Range
    
    Dim intProductId        As Integer
    Dim intCompanyId        As Integer
    Dim strProductName      As String
    Dim dblUnitPrice        As Double
    Dim strPackage          As String
    
    On Error GoTo CatchError
    
    Set rngValidationRange = shtData.Range("rProductID").CurrentRegion
    
    '// validate region size before continung
    If rngValidationRange.Rows.Count = 1 And rngValidationRange.Columns.Count Then
        GoTo CatchError
    End If
    
    Set rngValidationRange = rngValidationRange.Resize(rngValidationRange.CurrentRegion.Rows.Count - 1, rngValidationRange.Columns.Count)
    Set rngValidationRange = rngValidationRange.Offset(1, 0) '// offset to exclude headers
    
    '// check the validation first
    For Each rngMyCell In rngValidationRange
        If Not SheetDataValidation(rngMyCell) Then
            MsgBox "Please fix validation errors before import", vbInformation, "Validation Rules"
            Exit Sub
        End If
    Next rngMyCell
    
    '// then import the data to the database
    
    If Not InitializeConnection Then
        Exit Sub
    End If
    
    Set rngImport = shtData.Range("rProductID").CurrentRegion
    Set rngImport = rngValidationRange.Resize(rngImport.CurrentRegion.Rows.Count - 1, 1)
'    Set rngImport = rngImport.Offset(0, 1)

    '// loop round to get parameters
    For Each rngMyCell In rngImport
        
        intProductId = rngMyCell.Value
        '// clean the string for sql insert
        strProductName = rngMyCell.Offset(0, 2)
        strProductName = CleanForSQLInsert(strProductName)
        intCompanyId = shtControl.cboCompany.Value
        dblUnitPrice = rngMyCell.Offset(0, 3)
        strPackage = rngMyCell.Offset(0, 4)
        strPackage = CleanForSQLInsert(strPackage)
        
        If intProductId <> 0 Then
            '// use update proc
            m_strSQL = "EXEC spUpdateProducts " & _
                intProductId & ", " & _
                Chr(39) & strProductName & Chr(39) & ", " & _
                intCompanyId & ", " & _
                dblUnitPrice & ", " & _
                Chr(39) & strPackage & Chr(39)
        Else
            '// use insert proc
            m_strSQL = "EXEC spInsertProducts " & _
                Chr(39) & strProductName & Chr(39) & ", " & _
                intCompanyId & ", " & _
                dblUnitPrice & ", " & _
                Chr(39) & strPackage & Chr(39)
        End If
                
        Set m_rs = New ADODB.Recordset
        m_rs.Open m_strSQL, m_cn, adOpenStatic
        Set m_rs = Nothing
    Next rngMyCell
    
    '// refresh the data
    Call ImportData

ExitSub:
    Call TerminateConnection
    Exit Sub

CatchError:
        
    If rngValidationRange.Rows.Count = 1 And rngValidationRange.Columns.Count Then
        MsgBox "There does not appear to be any data to export." & vbNewLine & vbNewLine & _
            "Please select a company from the drop down box on the Control sheet.", vbInformation, "Validation Rule"
    Else
        MsgBox "Error Naumber: " & Err.Number & vbNewLine & vbNewLine & _
                "Error Message: " & Err.Description, vbCritical, "Run Time Error"
    End If
    
    GoTo ExitSub
    
End Sub





Sub RangeBorders(rngUsedRange As Range, blnClearBorders As Boolean)
    
    If Not blnClearBorders Then
        rngUsedRange.Borders(xlDiagonalDown).LineStyle = xlNone
        rngUsedRange.Borders(xlDiagonalUp).LineStyle = xlNone
        With rngUsedRange.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With rngUsedRange.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With rngUsedRange.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With rngUsedRange.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With rngUsedRange.Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With rngUsedRange.Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
    ElseIf blnClearBorders Then
        rngUsedRange.Borders(xlDiagonalDown).LineStyle = xlNone
        rngUsedRange.Borders(xlDiagonalUp).LineStyle = xlNone
        rngUsedRange.Borders(xlEdgeLeft).LineStyle = xlNone
        rngUsedRange.Borders(xlEdgeTop).LineStyle = xlNone
        rngUsedRange.Borders(xlEdgeBottom).LineStyle = xlNone
        rngUsedRange.Borders(xlEdgeRight).LineStyle = xlNone
        rngUsedRange.Borders(xlInsideVertical).LineStyle = xlNone
        rngUsedRange.Borders(xlInsideHorizontal).LineStyle = xlNone
    End If

End Sub


Function CleanForSQLInsert(strClean As String)

    CleanForSQLInsert = Replace(strClean, Chr(34), Chr(34) & Chr(34))
    CleanForSQLInsert = Replace(strClean, Chr(39), Chr(39) & Chr(39))

End Function

