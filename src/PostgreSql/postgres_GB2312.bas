Attribute VB_Name = "postgres"
Option Explicit


Private Const CP_UTF8 As Long = 65001
Private Const GB2312 As Long = 936
Private Const SYSENCONDING As Long = GB2312
'https://docs.microsoft.com/en-us/windows/win32/intl/code-page-identifiers
Private Const PGRES_COMMAND_OK As Integer = 2
Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long) As Long
Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName$) As LongPtr
'Connection
Private Declare PtrSafe Function PQconnectdb Lib "libpq" (ByVal conninfo$) As LongPtr 'connect pointer
    'postgresql://localhost
    'postgresql://localhost:5433
    'postgresql://localhost/mydb
    'postgresql://user@localhost
    'postgresql://user:secret@localhost
    'postgresql://other@localhost/otherdb?connect_timeout=10&application_name=myapp
    'postgresql://host1:123,host2:456/somedb?target_session_attrs=any&application_name=myapp
'Close
Private Declare PtrSafe Sub PQfinish Lib "libpq" (ByVal Connection As LongPtr)
'Status
Private Declare PtrSafe Function PQserverVersion Lib "libpq" (ByVal Connection As LongPtr) As LongPtr
Private Declare PtrSafe Function PQstatus Lib "libpq" (ByVal Connection As LongPtr) As LongPtr
Private Declare PtrSafe Function PQerrorMessage Lib "libpq" (ByVal Connection As LongPtr) As LongPtr
Private Declare PtrSafe Function PQresultErrorMessage Lib "libpq" (ByVal Result As LongPtr) As LongPtr
Private Declare PtrSafe Function PQresultStatus Lib "libpq" (ByVal Result As LongPtr) As LongPtr
'Exec
Private Declare PtrSafe Function PQexec Lib "libpq" (ByVal Connection As LongPtr, ByVal Command$) As LongPtr
'Private Declare PtrSafe Function PQprepare Lib "libpq" (connection As Long, ByVal stmtName$, ByVal query$, ByVal nParams As interger, ByVal paramTypers) As LongPtr
'Results
Private Declare PtrSafe Function PQnfields Lib "libpq" (ByVal Result As LongPtr) As Long
Private Declare PtrSafe Function PQntuples Lib "libpq" (ByVal Result As LongPtr) As Long
Private Declare PtrSafe Function PQfname Lib "libpq" (ByVal Result As LongPtr, ByVal column_number As Integer) As LongPtr
Private Declare PtrSafe Function PQftype Lib "libpq" (ByVal Result As LongPtr, ByVal column_number As Integer) As LongPtr
Private Declare PtrSafe Function PQgetvalue Lib "libpq" (ByVal Result As LongPtr, ByVal row_number As Integer, ByVal column_number As Integer) As LongPtr
Private Declare PtrSafe Sub PQclear Lib "libpq" (ByVal Result As LongPtr)
Dim postgresLibrary As LongPtr

Private Function PtrToString(ByVal pUtf8String As LongPtr, Optional ByVal Encodeing = SYSENCONDING) As String
    Dim buf As String
    Dim cSize As Long
    Dim retVal As Long
    
    cSize = MultiByteToWideChar(Encodeing, 0, pUtf8String, -1, 0, 0)
    ' cSize includes the terminating null character
    If cSize <= 1 Then
        PtrToString = vbNullString
        Exit Function
    End If
    
    PtrToString = String(cSize - 1, "*") ' and a termintating null char.
    retVal = MultiByteToWideChar(Encodeing, 0, pUtf8String, -1, StrPtr(PtrToString), cSize)
    If retVal = 0 Then
        Debug.Print "PtrToString Error:", Err.LastDllError
        Exit Function
    End If
End Function


Private Function postgres_Initialize(Optional ByVal libDir As String) As LongPtr
    If postgresLibrary = 0 Then
        If libDir = "" Then libDir = ThisWorkbook.Path
        If Right(libDir, 1) <> "\" Then libDir = libDir & "\"
        postgresLibrary = LoadLibrary(libDir + "libpq.dll")
        If postgresLibrary = 0 Then
            postgres_Initialize = 1
            Exit Function
        End If
    End If
    postgres_Initialize = 0
End Function

Public Function PSQLR(ByVal Connect$, ByVal Statement$, Optional ByVal Include_field_name As Boolean = False)
    On Error goto e
    Dim Connection As Longptr, res_pointer As Longptr
    Dim i As Long, j As Long, data$
    Dim Result As New vbaList, row As New vbaList
    If Connect = "" Then
        PSQLR = "Input connect uri"
        Exit Function
    End If
    If postgres_Initialize() = 1 Then
        PSQLR = "Missing libpq.dll"
        GoTo e
    End If
    Connection = PQconnectdb(Connect)
    If PQstatus(Connection) <> 0 Then
        PSQLR = PtrToString(PQerrorMessage(Connection), SYSENCONDING)
    Exit Function
    End If
    res_pointer = PQexec(Connection, Statement)
    If PQresultStatus(res_pointer) <> PGRES_COMMAND_OK Then
        PSQLR = Application.WorksheetFunction.Transpose(Split(PtrToString(PQresultErrorMessage(res_pointer), SYSENCONDING), Chr(10)))
        GoTo e
    End If

    If Include_field_name Then
        For i = 0 To PQnfields(res_pointer) - 1
            row.Add (PtrToString(PQfname(res_pointer, i)), SYSENCONDING)
        Next
        Result.Add (row.ToArray)
    End If

    For i = 0 To PQntuples(res_pointer) - 1
        row.RemoveAll
        For j = 0 To PQnfields(res_pointer) - 1
            data = PtrToString(PQgetvalue(res_pointer, i, j), SYSENCONDING)
            If IsNumeric(data) Then
                row.Add (Val(data))
            ElseIf IsDate(data) Then
                row.Add (DateValue(data))
            Else
                row.Add (data)
            End If
        Next
        Result.Add (row.ToArray)
    Next
    PSQLR = Result.ToArray
e:  PQclear (res_pointer)
    PQfinish Connection
End Function
