Attribute VB_Name = "sqlite3"
Option Explicit

' Notes:
' Microsoft uses UTF-16, little endian byte order.

Private Const JULIANDAY_OFFSET As Double = 2415018.5

' Returned from sqlite3Initialize
Private Const SQLITE_INIT_OK     As Long = 0
Private Const SQLITE_INIT_ERROR  As Long = 1

' SQLite data types
Private Const SQLITE_INTEGER  As Long = 1
Private Const SQLITE_FLOAT    As Long = 2
Private Const SQLITE_TEXT     As Long = 3
Private Const SQLITE_BLOB     As Long = 4
Private Const SQLITE_NULL     As Long = 5

' SQLite atandard return value
Private Const SQLITE_OK          As Long = 0   ' Successful result
Private Const SQLITE_ERROR       As Long = 1   ' SQL error or missing database
Private Const SQLITE_INTERNAL    As Long = 2   ' Internal logic error in SQLite
Private Const SQLITE_PERM        As Long = 3   ' Access permission denied
Private Const SQLITE_ABORT       As Long = 4   ' Callback routine requested an abort
Private Const SQLITE_BUSY        As Long = 5   ' The database file is locked
Private Const SQLITE_LOCKED      As Long = 6   ' A table in the database is locked
Private Const SQLITE_NOMEM       As Long = 7   ' A malloc() failed
Private Const SQLITE_READONLY    As Long = 8   ' Attempt to write a readonly database
Private Const SQLITE_INTERRUPT   As Long = 9   ' Operation terminated by sqlite3_interrupt()
Private Const SQLITE_IOERR      As Long = 10   ' Some kind of disk I/O error occurred
Private Const SQLITE_CORRUPT    As Long = 11   ' The database disk image is malformed
Private Const SQLITE_NOTFOUND   As Long = 12   ' NOT USED. Table or record not found
Private Const SQLITE_FULL       As Long = 13   ' Insertion failed because database is full
Private Const SQLITE_CANTOPEN   As Long = 14   ' Unable to open the database file
Private Const SQLITE_PROTOCOL   As Long = 15   ' NOT USED. Database lock protocol error
Private Const SQLITE_EMPTY      As Long = 16   ' Database is empty
Private Const SQLITE_SCHEMA     As Long = 17   ' The database schema changed
Private Const SQLITE_TOOBIG     As Long = 18   ' String or BLOB exceeds size limit
Private Const SQLITE_CONSTRAINT As Long = 19   ' Abort due to constraint violation
Private Const SQLITE_MISMATCH   As Long = 20   ' Data type mismatch
Private Const SQLITE_MISUSE     As Long = 21   ' Library used incorrectly
Private Const SQLITE_NOLFS      As Long = 22   ' Uses OS features not supported on host
Private Const SQLITE_AUTH       As Long = 23   ' Authorization denied
Private Const SQLITE_FORMAT     As Long = 24   ' Auxiliary database format error
Private Const SQLITE_RANGE      As Long = 25   ' 2nd parameter to sqlite3_bind out of range
Private Const SQLITE_NOTADB     As Long = 26   ' File opened that is not a database file
Private Const SQLITE_ROW        As Long = 100  ' sqlite3_step() has another row ready
Private Const SQLITE_DONE       As Long = 101  ' sqlite3_step() has finished executing

' Extended error codes
Private Const SQLITE_IOERR_READ               As Long = 266  '(SQLITE_IOERR | (1<<8))
Private Const SQLITE_IOERR_SHORT_READ         As Long = 522  '(SQLITE_IOERR | (2<<8))
Private Const SQLITE_IOERR_WRITE              As Long = 778  '(SQLITE_IOERR | (3<<8))
Private Const SQLITE_IOERR_FSYNC              As Long = 1034 '(SQLITE_IOERR | (4<<8))
Private Const SQLITE_IOERR_DIR_FSYNC          As Long = 1290 '(SQLITE_IOERR | (5<<8))
Private Const SQLITE_IOERR_TRUNCATE           As Long = 1546 '(SQLITE_IOERR | (6<<8))
Private Const SQLITE_IOERR_FSTAT              As Long = 1802 '(SQLITE_IOERR | (7<<8))
Private Const SQLITE_IOERR_UNLOCK             As Long = 2058 '(SQLITE_IOERR | (8<<8))
Private Const SQLITE_IOERR_RDLOCK             As Long = 2314 '(SQLITE_IOERR | (9<<8))
Private Const SQLITE_IOERR_DELETE             As Long = 2570 '(SQLITE_IOERR | (10<<8))
Private Const SQLITE_IOERR_BLOCKED            As Long = 2826 '(SQLITE_IOERR | (11<<8))
Private Const SQLITE_IOERR_NOMEM              As Long = 3082 '(SQLITE_IOERR | (12<<8))
Private Const SQLITE_IOERR_ACCESS             As Long = 3338 '(SQLITE_IOERR | (13<<8))
Private Const SQLITE_IOERR_CHECKRESERVEDLOCK  As Long = 3594 '(SQLITE_IOERR | (14<<8))
Private Const SQLITE_IOERR_LOCK               As Long = 3850 '(SQLITE_IOERR | (15<<8))
Private Const SQLITE_IOERR_CLOSE              As Long = 4106 '(SQLITE_IOERR | (16<<8))
Private Const SQLITE_IOERR_DIR_CLOSE          As Long = 4362 '(SQLITE_IOERR | (17<<8))
Private Const SQLITE_LOCKED_SHAREDCACHE       As Long = 265  '(SQLITE_LOCKED | (1<<8) )

' Flags For File Open Operations
Private Const SQLITE_OPEN_READONLY           As Long = 1       ' Ok for sqlite3_open_v2()
Private Const SQLITE_OPEN_READWRITE          As Long = 2       ' Ok for sqlite3_open_v2()
Private Const SQLITE_OPEN_CREATE             As Long = 4       ' Ok for sqlite3_open_v2()
Private Const SQLITE_OPEN_DELETEONCLOSE      As Long = 8       ' VFS only
Private Const SQLITE_OPEN_EXCLUSIVE          As Long = 16      ' VFS only
Private Const SQLITE_OPEN_AUTOPROXY          As Long = 32      ' VFS only
Private Const SQLITE_OPEN_URI                As Long = 64      ' Ok for sqlite3_open_v2()
Private Const SQLITE_OPEN_MEMORY             As Long = 128     ' Ok for sqlite3_open_v2()
Private Const SQLITE_OPEN_MAIN_DB            As Long = 256     ' VFS only
Private Const SQLITE_OPEN_TEMP_DB            As Long = 512     ' VFS only
Private Const SQLITE_OPEN_TRANSIENT_DB       As Long = 1024    ' VFS only
Private Const SQLITE_OPEN_MAIN_JOURNAL       As Long = 2048    ' VFS only
Private Const SQLITE_OPEN_TEMP_JOURNAL       As Long = 4096    ' VFS only
Private Const SQLITE_OPEN_SUBJOURNAL         As Long = 8192    ' VFS only
Private Const SQLITE_OPEN_MASTER_JOURNAL     As Long = 16384   ' VFS only
Private Const SQLITE_OPEN_NOMUTEX            As Long = 32768   ' Ok for sqlite3_open_v2()
Private Const SQLITE_OPEN_FULLMUTEX          As Long = 65536   ' Ok for sqlite3_open_v2()
Private Const SQLITE_OPEN_SHAREDCACHE        As Long = 131072  ' Ok for sqlite3_open_v2()
Private Const SQLITE_OPEN_PRIVATECACHE       As Long = 262144  ' Ok for sqlite3_open_v2()
Private Const SQLITE_OPEN_WAL                As Long = 524288  ' VFS only

' Options for Text and Blob binding
Private Const SQLITE_STATIC      As Long = 0
Private Const SQLITE_TRANSIENT   As Long = -1

' System calls
Private Const CP_UTF8 As Long = 65001
Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long) As Long
Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, ByVal lpDefaultChar As LongPtr, ByVal lpUsedDefaultChar As LongPtr) As Long
Private Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" (ByVal pDest As LongPtr, ByVal pSource As LongPtr, ByVal Length As Long)
Private Declare PtrSafe Function lstrcpynW Lib "kernel32" (ByVal pwsDest As LongPtr, ByVal pwsSource As LongPtr, ByVal cchCount As Long) As LongPtr
Private Declare PtrSafe Function lstrcpyW Lib "kernel32" (ByVal pwsDest As LongPtr, ByVal pwsSource As LongPtr) As LongPtr
Private Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal pwsString As LongPtr) As Long
Private Declare PtrSafe Function SysAllocString Lib "OleAut32" (ByRef pwsString As LongPtr) As LongPtr
Private Declare PtrSafe Function SysStringLen Lib "OleAut32" (ByVal bstrString As LongPtr) As Long
Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As LongPtr
Private Declare PtrSafe Function FreeLibrary Lib "kernel32" (ByVal hLibModule As LongPtr) As Long

'=====================================================================================

' SQLite library version
'Private Declare PtrSafe Function sqlite3_libversion Lib "sqlite3" () As LongPtr ' PtrUtf8String
' Database connections
Private Declare PtrSafe Function sqlite3_open16 Lib "sqlite3" (ByVal pwsFileName As LongPtr, ByRef hDb As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_open_v2 Lib "sqlite3" (ByVal pwsFileName As LongPtr, ByRef hDb As LongPtr, ByVal iFlags As Long, ByVal zVfs As LongPtr) As Long ' PtrDb
Private Declare PtrSafe Function sqlite3_close Lib "sqlite3" (ByVal hDb As LongPtr) As Long
' Database connection error info
Private Declare PtrSafe Function sqlite3_errmsg Lib "sqlite3" (ByVal hDb As LongPtr) As LongPtr ' PtrUtf8String

' Statements
Private Declare PtrSafe Function sqlite3_prepare16_v2 Lib "sqlite3" _
    (ByVal hDb As LongPtr, ByVal pwsSql As LongPtr, ByVal nSqlLength As Long, ByRef hStmt As LongPtr, ByVal ppwsTailOut As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_step Lib "sqlite3" (ByVal hStmt As LongPtr) As Long

' Statement column access (0-based indices)
Private Declare PtrSafe Function sqlite3_column_count Lib "sqlite3" (ByVal hStmt As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_column_type Lib "sqlite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As Long
Private Declare PtrSafe Function sqlite3_column_name Lib "sqlite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongPtr ' PtrString
Private Declare PtrSafe Function sqlite3_column_name16 Lib "sqlite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongPtr ' PtrWString
Private Declare PtrSafe Function sqlite3_column_blob Lib "sqlite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongPtr ' PtrData
Private Declare PtrSafe Function sqlite3_column_bytes Lib "sqlite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As Long
Private Declare PtrSafe Function sqlite3_column_bytes16 Lib "sqlite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As Long
Private Declare PtrSafe Function sqlite3_column_double Lib "sqlite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As Double
Private Declare PtrSafe Function sqlite3_column_int Lib "sqlite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As Long
Private Declare PtrSafe Function sqlite3_column_int64 Lib "sqlite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongLong
Private Declare PtrSafe Function sqlite3_column_text Lib "sqlite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongPtr ' PtrString
Private Declare PtrSafe Function sqlite3_column_text16 Lib "sqlite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongPtr ' PtrWString
' Private Declare PtrSafe Function sqlite3_column_value Lib "sqlite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongPtr ' Ptrsqlite3Value

' Initialize - load libraries explicitly
Private hSQLiteLibrary As LongPtr
Private hSQLiteStdCallLibrary As LongPtr

Private Function sqlite3_Initialize(Optional ByVal libDir As String) As Long
    ' A nice option here is to call SetDllDirectory, but that API is only available since Windows XP SP1.
    If hSQLiteLibrary = 0 Then
        If libDir = "" Then libDir = ThisWorkbook.Path
        If Right(libDir, 1) <> "\" Then libDir = libDir & "\"
        hSQLiteLibrary = LoadLibrary(libDir + "sqlite3.dll")
        If hSQLiteLibrary = 0 Then
            Debug.Print "sqlite3_Initialize Error Loading " + libDir + "sqlite3.dll:", Err.LastDllError
            sqlite3_Initialize = SQLITE_INIT_ERROR
            Exit Function
        End If
    End If
    sqlite3_Initialize = SQLITE_INIT_OK
End Function

Private Sub sqlite3Free()
   Dim refCount As Long
   If hSQLiteStdCallLibrary <> 0 Then
        refCount = FreeLibrary(hSQLiteStdCallLibrary)
        hSQLiteStdCallLibrary = 0
        If refCount = 0 Then
            Debug.Print "sqlite3Free Error Freeing sqlite3_StdCall.dll:", refCount, Err.LastDllError
        End If
    End If
    If hSQLiteLibrary <> 0 Then
        refCount = FreeLibrary(hSQLiteLibrary)
        hSQLiteLibrary = 0
        If refCount = 0 Then
            Debug.Print "sqlite3Free Error Freeing sqlite3.dll:", refCount, Err.LastDllError
        End If
    End If
End Sub

'=====================================================================================
' Statement column access (0-based indices)


Private Function sqlite3ColumnName(ByVal stmtHandle As LongPtr, ByVal ZeroBasedColIndex As Long) As String
    sqlite3ColumnName = Utf8PtrToString(sqlite3_column_name(stmtHandle, ZeroBasedColIndex))
End Function


Private Function sqlite3ColumnText(ByVal stmtHandle As LongPtr, ByVal ZeroBasedColIndex As Long) As String
    sqlite3ColumnText = Utf8PtrToString(sqlite3_column_text(stmtHandle, ZeroBasedColIndex))
End Function


Private Function sqlite3ColumnDate(ByVal stmtHandle As LongPtr, ByVal ZeroBasedColIndex As Long) As Date
    sqlite3ColumnDate = FromJulianDay(sqlite3_column_double(stmtHandle, ZeroBasedColIndex))
End Function

Private Function sqlite3ColumnBlob(ByVal stmtHandle As LongPtr, ByVal ZeroBasedColIndex As Long) As Byte()
    Dim ptr As LongPtr
    Dim Length As Long
    Dim buf() As Byte
    
    ptr = sqlite3_column_blob(stmtHandle, ZeroBasedColIndex)
    Length = sqlite3_column_bytes(stmtHandle, ZeroBasedColIndex)
    ReDim buf(Length - 1)
    RtlMoveMemory VarPtr(buf(0)), ptr, Length
    sqlite3ColumnBlob = buf
End Function

' String Helpers
Private Function Utf8PtrToString(ByVal pUtf8String As LongPtr) As String
    Dim buf As String
    Dim cSize As Long
    Dim retVal As Long
    
    cSize = MultiByteToWideChar(CP_UTF8, 0, pUtf8String, -1, 0, 0)
    ' cSize includes the terminating null character
    If cSize <= 1 Then
        Utf8PtrToString = ""
        Exit Function
    End If
    
    Utf8PtrToString = String(cSize - 1, "*") ' and a termintating null char.
    retVal = MultiByteToWideChar(CP_UTF8, 0, pUtf8String, -1, StrPtr(Utf8PtrToString), cSize)
    If retVal = 0 Then
        Debug.Print "Utf8PtrToString Error:", Err.LastDllError
        Exit Function
    End If
End Function

Private Function StringToUtf8Bytes(ByVal str As String) As Variant
    Dim bSize As Long
    Dim retVal As Long
    Dim buf() As Byte
    
    bSize = WideCharToMultiByte(CP_UTF8, 0, StrPtr(str), -1, 0, 0, 0, 0)
    If bSize = 0 Then
        Exit Function
    End If
    
    ReDim buf(bSize)
    retVal = WideCharToMultiByte(CP_UTF8, 0, StrPtr(str), -1, VarPtr(buf(0)), bSize, 0, 0)
    If retVal = 0 Then
        Debug.Print "StringToUtf8Bytes Error:", Err.LastDllError
        Exit Function
    End If
    StringToUtf8Bytes = buf
End Function

Private Function Utf16PtrToString(ByVal pUtf16String As LongPtr) As String
    Dim StrLen As Long
    StrLen = lstrlenW(pUtf16String)
    Utf16PtrToString = String(StrLen, "*")
    lstrcpynW StrPtr(Utf16PtrToString), pUtf16String, StrLen
End Function

' Date Helpers
Private Function ToJulianDay(oleDate As Date) As Double
    ToJulianDay = CDbl(oleDate) + JULIANDAY_OFFSET
End Function

Private Function FromJulianDay(julianDay As Double) As Date
    FromJulianDay = CDate(julianDay - JULIANDAY_OFFSET)
End Function

Private Function ColumnValue(ByVal stmtHandle As LongPtr, ByVal ZeroBasedColIndex As Long, ByVal SQLiteType As Long) As Variant
    Select Case SQLiteType
        Case SQLITE_INTEGER:
            ColumnValue = sqlite3_column_int(stmtHandle, ZeroBasedColIndex)
        Case SQLITE_FLOAT:
            ColumnValue = sqlite3_column_double(stmtHandle, ZeroBasedColIndex)
        Case SQLITE_TEXT:
            Dim t$
            t = sqlite3ColumnText(stmtHandle, ZeroBasedColIndex)
            If IsDate(t) Then
                ColumnValue = DateValue(t)
            Else:
                ColumnValue = t
            End If
            'ColumnValue = sqlite3ColumnText(stmtHandle, ZeroBasedColIndex)
        Case SQLITE_BLOB:
            ColumnValue = sqlite3ColumnBlob(stmtHandle, ZeroBasedColIndex)
        Case SQLITE_NULL:
            ColumnValue = vbNullString
    End Select
End Function

Public Function SQLR(ByVal dbfile$, ByVal Statements$, Optional ByVal Include_field_name As Boolean = False)
    Dim myDbHandle As LongPtr, stmtHandle As LongPtr, statement$
    Dim colCount As Long, colType As Long, i%, j%, sqlstatus%
    Dim Result As New vbaList, col As New vbaList
    If dbfile = "" Then
        SQLR = "Input sqlite db"
        Exit Function
    End If
    If sqlite3_Initialize() = 1 Then SQLR = "Missing sqlite3.dll"
    If Dir(dbfile, 16) = Empty Then
        SQLR = "Database not exists!"
        Exit Function
    End If
    sqlite3_open16 StrPtr(dbfile), myDbHandle

    For Each statement In Split(Statements, ";")
        If statement = "" Then GoTo em
        sqlite3_prepare16_v2 myDbHandle, StrPtr(statement), Len(statement) * 2, stmtHandle, 0
        Do
            sqlstatus = sqlite3_step(stmtHandle)
            If sqlstatus = SQLITE_DONE Then
                Exit Do
            ElseIf sqlstatus = SQLITE_ROW Then
                GoTo pa
            Else
                SQLR = Utf8PtrToString(sqlite3_errmsg(myDbHandle))
                Exit Function
            End If
pa:         colCount = sqlite3_column_count(stmtHandle)
            If j = 0 And Include_field_name Then
                For i = 0 To colCount - 1
                    col.Add (sqlite3ColumnName(stmtHandle, i))
                Next
                Result.Add (col.ToArray)
                col.RemoveAll
                j = j + 1
            End If
            For i = 0 To colCount - 1
                colType = sqlite3_column_type(stmtHandle, i)
                col.Add (ColumnValue(stmtHandle, i, colType))
            Next
            Result.Add (col.ToArray)
            col.RemoveAll
        Loop
    Next
em: SQLR = Result.ToArray
    sqlite3_close myDbHandle
End Function
