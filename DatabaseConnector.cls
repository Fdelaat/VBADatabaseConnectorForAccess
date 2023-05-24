VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DatabaseConnector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule ParameterCanBeByVal, VariableNotAssigned, VariableNotUsed, UnassignedVariableUsage, ModuleWithoutFolder, IntegerDataType, HungarianNotation, UseMeaningfulName, EmptyIfBlock, FunctionReturnValueAlwaysDiscarded
Option Explicit

'Database connector V2.5  © Frank de Laat 11-05-2023

Const PRVD As String = "Microsoft.ace.OLEDB.12.0;"  'This is the connection provider.
Const DBPATH As String = "PATH TO THE ACCESS '.accdb' file HERE"
Const PWD As String = "DB PASSWORD HERE"
Const CONSTRING As String = "Provider=" & PRVD & "Data Source=" & DBPATH & ";Persist Security Info=False;Jet OLEDB:Database Password=" & PWD  'Info connectionstrings.com

Dim conn As New ADODB.Connection
Dim cmd As New ADODB.Command

Public Sub InsertNew(ByRef dbTable As String, ByRef dbFields As Variant, ByRef dbFieldValues As Variant)
'Create
On Error GoTo ERR_ADODBINS:

Dim strFieldName As String
Dim strFieldValue As String
    
    strFieldName = ""
    strFieldValue = ""
    
    'Array to stringconversion by join and StrValueConverter methode
    'works currently only with a single record due to array limitation
    strFieldName = Join(dbFields, ",")
    strFieldValue = StrValueConverter(dbFieldValues)
    
    'This is the connection string that you will require when opening the the connection.
    conn.ConnectionString = CONSTRING
    
    'Opening the connection
    conn.CursorLocation = adUseClient
    conn.Open
    
    'Running the query on the open connection. It will send the data in the Database.
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    cmd.CommandText = "INSERT INTO " & dbTable & " (" & strFieldName & ") VALUES (" & strFieldValue & ");"

    cmd.Execute , , adExecuteNoRecords
    
    'Closing the connections
    conn.Close


Exit Sub

ERR_ADODBINS:

    Select Case Err
    
        Case Else:
            MsgBox "Database connectie bevat fouten!" & vbCrLf & "Foutcode: " & Err.Number & vbCrLf & Err.Description & "", vbOKOnly, "Database Connector - InsertNew!"
    
    End Select
End Sub

Public Function SqlQuery(ByRef query As String, ByRef blConvertToRow As Boolean, Optional ByRef strFieldSortname As String, Optional ByRef strFieldSortDirection As String) As Variant
' Read :This function returns array of requested SQL Query input needs to be formated "SELECT..FROM..;"
On Error GoTo ERR_ADODBSQ:

'Creating objects of Connection and Recordset
Dim arrData As Variant
Dim arrRes As Variant
Dim rec As ADODB.Recordset
Dim X As Integer
Dim Y As Integer

Set rec = New ADODB.Recordset

    'This is the connection string that you will require when opening the the connection. Info connectionstrings.com
    conn.ConnectionString = CONSTRING
    
    'Opening the connection
    conn.CursorLocation = adUseClient
    conn.Open
    
    'Running the query on the open connection. It will get all the data in the rec object.
    Set rec.ActiveConnection = conn
    rec.LockType = adLockOptimistic
    rec.CursorLocation = adUseClient
    rec.Source = query
    rec.Open
    
    'Added  Optional Sorting method for 1 Field an Direction (ASC/DESC) d.d.14-02-2022
    If strFieldSortname = "" Or strFieldSortDirection = "" Then
    Else
        rec.Sort = strFieldSortname & " " & Format$(strFieldSortDirection, ">")
    End If
    
    'If rec has no value return empty 2-d array
    If Not rec.EOF = True Then
            'Call ConvertRstToArr function in this module
            arrData = ConvertRstToArr(rec)
        
        'Closing the connections
        rec.Close
        conn.Close
        
        'Remove null for array
        For X = 0 To UBound(arrData, 1)
            
            For Y = 0 To UBound(arrData, 2)
                
                If VarType(arrData(X, Y)) = 1 Then
                    arrData(X, Y) = VarType(12)
                    arrData(X, Y) = vbNullString
                End If
            
            Next
            
        Next
        
        'Rebuild array to columns per row when blAsRow = THRUE
        X = 0
        Y = 0
        
        If blConvertToRow = True Then
            
            ReDim arrRes(0 To UBound(arrData, 2), 0 To UBound(arrData, 1))
            
            For X = 0 To UBound(arrData, 1)
                
                For Y = 0 To UBound(arrData, 2)
                    arrRes(Y, X) = arrData(X, Y)
                Next
                
            Next
            
            SqlQuery = arrRes
            
        Else
        
        'Return the array if blAsRow = FALSE
        SqlQuery = arrData
        
        End If
    Else
    
        'Closing the connections
        rec.Close
        conn.Close
        
        ReDim arrData(0 To 0, 0 To 0)
        arrData(0, 0) = ""
        SqlQuery = arrData
    
    End If

Exit Function

ERR_ADODBSQ:

    Select Case Err
    
        Case -2147217843:
            MsgBox "Het Database wachtwoord is niet juist!" & vbCrLf & "Foutcode: " & Err.Number & vbCrLf & Err.Description & "", vbOKOnly, "Database Connector - SqlQuery!"
        
        Case 3021:
            MsgBox "De QUERY bevat fouten!" & vbCrLf & "Foutcode: " & Err.Number & vbCrLf & Err.Description & "", vbOKOnly, "Database Connector - SqlQuery!"
    
        Case Else:
            MsgBox "Database connectie bevat fouten!" & vbCrLf & "Foutcode: " & Err.Number & vbCrLf & Err.Description & "", vbOKOnly, "Database Connector - SqlQuery!"
    
    End Select
End Function

Public Sub Update(ByRef dbTable As String, ByRef dbFields As Variant, ByRef dbFieldValues As Variant, ByRef dbRecordNr As LongPtr)
'Update
On Error GoTo ERR_ADODBUPD:

Dim strConcat As String
Dim strFieldName As String
Dim strFieldValue As String
Dim i As Integer

    'Array to stringconversion by join and StrValueConverter methode
    strFieldName = Join(dbFields, ",")
    strFieldValue = StrValueConverter(dbFieldValues)
    
    'Build the SQL update string in this patern "field equals value, nextfield equals nextvalue, endfield equals endvalue"
    'works only with a single record due to limitation MS Access
    If Not UBound(dbFields) = 0 Then 'If there is only 1 field the trailing comma needs to be removed
        For i = 0 To UBound(dbFields)
            If i = 0 Then
                strConcat = dbFields(i) & Chr$(32) & Chr$(61) & Chr$(32) & dbFieldValues(i) & Chr$(44)
            Else
                If Not i = UBound(dbFields) Then
                    strConcat = strConcat & dbFields(i) & Chr$(32) & Chr$(61) & Chr$(32) & dbFieldValues(i) & Chr$(44)
                Else
                    strConcat = strConcat & dbFields(i) & Chr$(32) & Chr$(61) & Chr$(32) & dbFieldValues(i) 'Todo: If Last value (acdb:numeric datatype) is empty  datatype error is given due to 'text' datatype solution needed
                End If
            End If
        Next
    Else
        strConcat = dbFields(i) & Chr$(32) & Chr$(61) & Chr$(32) & dbFieldValues(i)
    End If
    
    'This is the connection string that you will require when opening the the connection. Info connectionstrings.com
    conn.ConnectionString = CONSTRING
    
    'Opening the connection
    conn.CursorLocation = adUseClient
    conn.Open
    
    'Running the query on the open connection. It will update the data in the Database.
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    cmd.CommandText = "UPDATE " & dbTable & " SET " & strConcat & " WHERE Id=" & dbRecordNr & ";"
    
    cmd.Execute , , adExecuteNoRecords
    
    'Closing the connections
    conn.Close
    
Exit Sub

ERR_ADODBUPD:

    Select Case Err
    
        Case Else:
            MsgBox "Database connectie bevat fouten!" & vbCrLf & "Foutcode: " & Err.Number & vbCrLf & Err.Description & "", vbOKOnly, "Database Connector - Update!"
    
    End Select
End Sub

Public Sub Delete(ByRef dbTable As String, ByRef dbRecordNr As LongPtr)
'Destroy
On Error GoTo ERR_ADODBDEL:

    'This is the connection string that you will require when opening the the connection. Info connectionstrings.com
    conn.ConnectionString = CONSTRING
    
    'Opening the connection
    conn.CursorLocation = adUseClient
    conn.Open
    
    'Running the query on the open connection. It will delete the data in the Database.
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    cmd.CommandText = "DELETE * FROM " & dbTable & " WHERE Id=" & dbRecordNr & ";"
    cmd.Execute , , adExecuteNoRecords
    
    'Closing the connections
    conn.Close

Exit Sub

ERR_ADODBDEL:

    Select Case Err
    
        Case Else:
            MsgBox "Database connectie bevat fouten!" & vbCrLf & "Foutcode: " & Err.Number & vbCrLf & Err.Description & "", vbOKOnly, "Database Connector - Delete!"
    
    End Select
End Sub


Private Function ConvertRstToArr(ByRef Summary As Object) As Variant 'Input need to be a recordset

    ConvertRstToArr = Summary.GetRows
    
End Function

Private Function StrValueConverter(ByRef dbFieldValues As Variant) As String

Dim strFieldValue As String
Dim i As Integer

    strFieldValue = ""
    
    'Add delimiter Chr(39) to values excluding booleans
    For i = LBound(dbFieldValues) To UBound(dbFieldValues)
               
        Select Case VarType(dbFieldValues(i))
        
        Case 5:
            dbFieldValues(i) = DblToStr(dbFieldValues(i))
            
        Case 7:
            dbFieldValues(i) = Chr$(35) & dbFieldValues(i) & Chr$(35)
            
        Case 11:
            'Translating issue set boolean to english verb
            If dbFieldValues(i) = True Then
                dbFieldValues(i) = "True"
            Else
                dbFieldValues(i) = "False"
            End If
        
        Case Else:
            ' Add leading and trailing single qoute and escape single qoute char(s) in dbFieldValues
            dbFieldValues(i) = Chr$(39) & Replace(CStr(dbFieldValues(i)), Chr$(39), Chr$(39) & Chr$(39), , , vbTextCompare) & Chr$(39)
        
        End Select
        
    Next
        
    'Convert array to stringlist
    strFieldValue = Join(dbFieldValues, ",")
    
    StrValueConverter = strFieldValue

End Function

Private Function DblToStr(ByRef X As Variant)
 DblToStr = CStr(X)

 If (Application.ThousandsSeparator = ".") Then
  DblToStr = Replace(DblToStr, ".", "")
 End If
 If (Application.DecimalSeparator = ",") Then
  DblToStr = Replace(DblToStr, ",", ".")
 End If
  
End Function

Private Sub Class_Terminate()
    Set conn = Nothing
    Set cmd = Nothing
End Sub
