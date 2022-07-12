Attribute VB_Name = "mysql"
Option Explicit

Private Const CP_UTF8 As Long = 65001
Private Const GB2312 As Long = 936
'https://docs.microsoft.com/en-us/windows/win32/intl/code-page-identifiers
Private Const LONGPTR_SIZE = 8
Private Const BYTE_SIZE = 1

'kernel32 lib
Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long) As Long
Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName$) As LongPtr
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDestination As Any, lpSource As Any, ByVal lLength As Long)

' connection
Private Declare PtrSafe Function mysql_real_connect Lib "libmySQL" (ByVal handler As LongPtr, ByVal host$, ByVal user$, ByVal password$, ByVal dbname$, ByVal port%, Optional ByVal socket = Null, Optional ByVal flag% = 0) As LongPtr
Private Declare PtrSafe Function mysql_init Lib "libmySQL" (Optional ByVal init As LongPtr) As LongPtr
'user:secret@localhost:port/database
'opt
Private Declare PtrSafe Function mysql_set_character_set Lib "libmySQL" (ByVal lMysql As LongPtr, ByVal cs_name As String) As LongPtr

'Close
Private Declare PtrSafe Sub mysql_close Lib "libmySQL" (ByVal Connection As LongPtr)
'Status
Private Declare PtrSafe Function mysql_error Lib "libmySQL" (ByVal Connection As LongPtr) As LongPtr
'Private Declare PtrSafe Function mysql_get_server_version Lib "libmySQL" (ByVal Connection As LongPtr) As Long
'Exec
Private Declare PtrSafe Function mysql_query Lib "libmySQL" (ByVal handler As LongPtr, ByVal Command$) As LongPtr

'Results
Private Declare PtrSafe Function mysql_store_result Lib "libmySQL" (ByVal Result As LongPtr) As LongPtr
Private Declare PtrSafe Function mysql_num_fields Lib "libmySQL" (ByVal Result As LongPtr) As Long
Private Declare PtrSafe Function mysql_fetch_lengths Lib "libmySQL" (ByVal Result As LongPtr) As LongPtr
Private Declare PtrSafe Function mysql_fetch_row Lib "libmySQL" (ByVal Result As LongPtr) As LongPtr 'arrary pointer
Private Declare PtrSafe Function mysql_fetch_field_direct Lib "libmySQL" (ByVal Result As LongPtr, ByVal lFieldNum As Integer) As LongPtr
'MYSQL Field Struct
Public Type typ_MYSQL_FIELD
    FieldName As LongPtr 
    org_Name As LongPtr 
    Table As LongPtr 
    org_table As LongPtr 
    Db As LongPtr
    Catalog As LongPtr 
    Def As LongPtr 
    Length As Long 
    Max_length As Long
    Name_length As Integer
    org_name_length As Integer
    Table_length As Integer
    org_table_length As Integer
    Db_length As Integer
    Catalog_length As Integer
    Def_length As Integer
    Flags As Integer
    Decimals As Integer
    Charsetnr As Integer
    FieldType As Long
    'extension As LongPtr 'void pointer
End Type

Dim mysqlLibrary As LongPtr

Private Function PtrToString(ByVal pUtf8String As LongPtr, Optional ByVal Encodeing$ = CP_UTF8) As String

    Dim buf As String
    Dim cSize As Long
    Dim retVal As Long
    
    cSize = MultiByteToWideChar(Encodeing, 0, pUtf8String, -1, 0, 0)
    ' cSize includes the terminating null character
    If cSize <= 1 Then
        PtrToString = ""
        Exit Function
    End If
    
    PtrToString = String(cSize - 1, "*") ' and a termintating null char.
    retVal = MultiByteToWideChar(Encodeing, 0, pUtf8String, -1, StrPtr(PtrToString), cSize)
    If retVal = 0 Then
        Debug.Print "PtrToString Error:", Err.LastDllError
        Exit Function
    End If
End Function

Private Function mysqlInitialize(Optional ByVal libDir As String) As LongPtr
    If mysqlLibrary = 0 Then
        If libDir = "" Then libDir = ThisWorkbook.Path
        If Right(libDir, 1) <> "\" Then libDir = libDir & "\"
        mysqlLibrary = LoadLibrary(libDir + "libmysql.dll")
        If mysqlLibrary = 0 Then
            mysqlInitialize = 1
            Exit Function
        End If
    End If
    mysqlInitialize = 0
End Function


Private Function GetFieldName(ByVal SQL_RES As LongPtr, ByVal lField As Long)
    Dim lMYSQL_FIELD As LongPtr, a, b
    Dim mField As typ_MYSQL_FIELD
    If lField < 0 Then
        Exit Function
    End If
    lMYSQL_FIELD = mysql_fetch_field_direct(SQL_RES, lField)
    If lMYSQL_FIELD = 0 Then Exit Function
    CopyMemory mField, ByVal lMYSQL_FIELD, LenB(mField)
    GetFieldName = PtrToString(mField.FieldName, GB2312)
End Function

Private Function Get_Value(ByVal m_MYSQL_ROW As LongPtr, ByVal numlens As LongPtr, ByVal numfeilds As Long) As Variant
    On Error Resume Next
    Dim lRowData As LongPtr, data As String
    'get pointer to requested field
    CopyMemory lRowData, ByVal (m_MYSQL_ROW + (LONGPTR_SIZE * numfeilds)), LONGPTR_SIZE
    If lRowData = 0 Then
        Get_Value = vbNullString
    Else
        Get_Value = PtrToString(lRowData, GB2312)
    End If
End Function

Private Function uri_connect(ByVal handler As LongPtr, ByVal uri$) As LongPtr
    'user:secret@localhost:port/database
    Dim a$(), b$(), c$()
    a = Split(uri, "/") 'user:secret@localhost:port,database
    b = Split(a(0), ":") 'user,secret@localhost,port
    c = Split(b(1), "@") 'secret,localhost
    uri_connect = mysql_real_connect(handler, c(1), b(0), c(0), a(1), b(2))
End Function

Public Function MSQLR(ByVal Connect$, ByVal Statement$, Optional ByVal include_field_name As Boolean = False)
    On Error GoTo e
    Dim res_pointer As LongPtr, mysqlHandler As LongPtr, single_row As LongPtr
    Dim i As Long, lens As LongPtr, cols As Long, data$
    Dim Result As New vbaList, row As New vbaList
    If Connect = "" Then
        MSQLR = "Input connect uri"
        Exit Function
    End If
    If mysqlInitialize() = 1 Then
        MSQLR = "Missing libmysql.dll"
        GoTo e
    End If
    mysqlHandler = mysql_init(mysqlHandler)
    uri_connect mysqlHandler, Connect
    mysql_set_character_set mysqlHandler, "gb2312"
    If mysql_query(mysqlHandler, Statement) <> 0 Then
        MSQLR = PtrToString(mysql_error(mysqlHandler), GB2312)
        GoTo e
    End If
    res_pointer = mysql_store_result(mysqlHandler)
    cols = mysql_num_fields(res_pointer)
    'column
    If include_field_name Then
        For i = 0 To cols - 1
            row.Add (GetFieldName(res_pointer, i))
        Next
        Result.Add (row.ToArray)
    End If
    Do
        row.RemoveAll
        single_row = mysql_fetch_row(res_pointer)
        If single_row = 0 Then Exit Do
        For i = 0 To cols - 1
            data = Get_Value(single_row, mysql_fetch_lengths(res_pointer), i)
            If IsNumeric(data) Then
                row.Add (Val(data))
            ElseIf IsDate(data) Then
                row.Add (CDate(data))
            Else
                row.Add (data)
            End If
        Next
        Result.Add (row.ToArray)
    Loop
    MSQLR = Result.ToArray
e:  mysql_close (mysqlHandler)
End Function
