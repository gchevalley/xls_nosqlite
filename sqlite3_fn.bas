Attribute VB_Name = "sqlite3_fn"
Function SQLite3ConnectToDB(ByVal dbFile As String)

Dim dbHandle As Long
Dim retValue As Long

retValue = SQLite3Open(dbFile, dbHandle)
SQLite3ConnectToDB = dbHandle

End Function


Sub PrintColumns(ByVal stmtHandle As Long)
    Dim colCount As Long
    Dim colName As String
    Dim colType As Long
    Dim colTypeName As String
    Dim colValue As Variant
    
    Dim i As Long
    
    colCount = SQLite3ColumnCount(stmtHandle)
    Debug.Print "Column count: " & colCount
    For i = 0 To colCount - 1
        colName = SQLite3ColumnName(stmtHandle, i)
        colType = SQLite3ColumnType(stmtHandle, i)
        colTypeName = sqlite_TypeName(colType)
        colValue = ColumnValue(stmtHandle, i, colType)
        Debug.Print "Column " & i & ":", colName, colTypeName, colValue
    Next
End Sub

Sub PrintParameters(ByVal stmtHandle As Long)
    Dim paramCount As Long
    Dim paramName As String
    
    Dim i As Long
    
    paramCount = SQLite3BindParameterCount(stmtHandle)
    Debug.Print "Parameter count: " & paramCount
    For i = 1 To paramCount
        paramName = SQLite3BindParameterName(stmtHandle, i)
        Debug.Print "Parameter " & i & ":", paramName
    Next
End Sub


Function sqlite_TypeName(ByVal SQLiteType As Long) As String
    Select Case SQLiteType
        Case SQLITE_INTEGER:
            TypeName = "INTEGER"
        Case SQLITE_FLOAT:
            TypeName = "FLOAT"
        Case SQLITE_TEXT:
            TypeName = "TEXT"
        Case SQLITE_BLOB:
            TypeName = "BLOB"
        Case SQLITE_NULL:
            TypeName = "NULL"
    End Select
End Function

Function ColumnValue(ByVal stmtHandle As Long, ByVal ZeroBasedColIndex As Long, ByVal SQLiteType As Long) As Variant
    Select Case SQLiteType
        Case SQLITE_INTEGER:
            ColumnValue = SQLite3ColumnInt32(stmtHandle, ZeroBasedColIndex)
        Case SQLITE_FLOAT:
            ColumnValue = SQLite3ColumnDouble(stmtHandle, ZeroBasedColIndex)
        Case SQLITE_TEXT:
            ColumnValue = SQLite3ColumnText(stmtHandle, ZeroBasedColIndex)
        Case SQLITE_BLOB:
            ColumnValue = SQLite3ColumnText(stmtHandle, ZeroBasedColIndex)
        Case SQLITE_NULL:
            ColumnValue = Null
    End Select
End Function

' SQLite3 Helper Functions
Public Function SQLite3ExecuteNonQuery(ByVal dbHandle As Long, ByVal SqlCommand As String) As Long
    Dim stmtHandle As Long
    
    SQLite3PrepareV2 dbHandle, SqlCommand, stmtHandle
    SQLite3Step stmtHandle
    SQLite3Finalize stmtHandle
    
    SQLite3ExecuteNonQuery = SQLite3Changes(dbHandle)
End Function

Public Sub SQLite3ExecuteQuery(ByVal dbHandle As Long, ByVal sqlQuery As String)
    ' Dumps a query to the debug window. No error checking
    
    Dim stmtHandle As Long
    Dim RetVal As Long

    RetVal = SQLite3PrepareV2(dbHandle, sqlQuery, stmtHandle)
    Debug.Print "SQLite3PrepareV2 returned " & RetVal
    
    ' Start running the statement
    RetVal = SQLite3Step(stmtHandle)
    If RetVal = SQLITE_ROW Then
        Debug.Print "SQLite3Step Row Ready"
        PrintColumns stmtHandle
    Else
        Debug.Print "SQLite3Step returned " & RetVal
    End If
    
    ' Move to next row
    RetVal = SQLite3Step(stmtHandle)
    Do While RetVal = SQLITE_ROW
        Debug.Print "SQLite3Step Row Ready"
        PrintColumns stmtHandle
        RetVal = SQLite3Step(stmtHandle)
    Loop

    If RetVal = SQLITE_DONE Then
        Debug.Print "SQLite3Step Done"
    Else
        Debug.Print "SQLite3Step returned " & RetVal
    End If
    
    ' Finalize (delete) the statement
    RetVal = SQLite3Finalize(stmtHandle)
    Debug.Print "SQLite3Finalize returned " & RetVal
End Sub


Function sqlite3_get_field_type(ByVal local_type As Variant) As String

If UCase(Left(local_type, 3)) = "INT" Or UCase(Left(local_type, 4)) = "BOOL" Then
    sqlite3_get_field_type = "INTEGER"
ElseIf UCase(Left(local_type, 3)) = "TXT" Or UCase(Left(local_type, 4)) = "TEXT" Or UCase(Left(local_type, 3)) = "VAR" Then
    sqlite3_get_field_type = "TEXT"
ElseIf UCase(Left(local_type, 3)) = "DBL" Or UCase(Left(local_type, 6)) = "DOUBLE" Or UCase(Left(local_type, 4)) = "REAL" Then
    sqlite3_get_field_type = "REAL"
ElseIf UCase(Left(local_type, 3)) = "NUM" Or UCase(Left(local_type, 7)) = "NUMERIC" Or UCase(Left(local_type, 4)) = "DATE" Then
    sqlite3_get_field_type = "NUMERIC"
ElseIf UCase(Left(local_type, 4)) = "BLOB" Then
    sqlite3_get_field_type = "BLOB"
Else
    MsgBox ("error type")
End If

End Function


Function sqlite3_create_db(ByVal db_path As String, Optional ByVal delete_if_exists As Boolean = False) As Long 'retourne le dbHandle

Dim i As Double, j As Double, k As Double, m As Double, n As Double

Dim src_path As String
src_path = StrReverse(Mid(StrReverse(ThisWorkbook.FullName), InStr(StrReverse(ThisWorkbook.FullName), "\")))


Dim initReturn As Long
Dim dbHandle As Long
Dim stmHandle As Long
Dim retValue As Long


'initReturn = SQLite3Initialize(src_path & folder_dlls)
'
'retValue = SQLite3Open(DB_Path, dbHandle)
'sqlite3_create_db = dbHandle


If exist_file(db_path) = False Or delete_if_exists = True Then
    initReturn = SQLite3Initialize(src_path & folder_dlls)

    retValue = SQLite3Open(db_path, dbHandle)
    sqlite3_create_db = dbHandle
Else

    sqlite3_create_db = -1
    Exit Function
End If


End Function



Function sqlite3_create_tables(ByVal db_path As Variant, ByVal sqlite3_vec_sql_stmt_create_table As Variant, Optional drop_if_already_exists As Boolean = False) As Boolean

Dim i As Double, j As Double, k As Double, m As Double, n As Double

Dim src_path As String
src_path = StrReverse(Mid(StrReverse(ThisWorkbook.FullName), InStr(StrReverse(ThisWorkbook.FullName), "\")))


Dim initReturn As Long
Dim dbHandle As Long
Dim stmHandle As Long
Dim retValue As Long

If IsNumeric(db_path) Then
    dbHandle = db_path
Else
    SQLite3Initialize src_path & folder_dlls
    SQLite3Open db_path, dbHandle
    SQLite3Finalize stmHandle
End If

sqlite3_create_tables = True

Dim tmp_table_name
Dim need_to_create_db As Boolean
need_to_create_db = True

For i = 0 To UBound(sqlite3_vec_sql_stmt_create_table, 1)

    's'assure que n'existe pas deja
    tmp_table_name = Mid(sqlite3_vec_sql_stmt_create_table(i), Len("CREATE TABLE ") + 1, InStr(Len("CREATE TABLE ") + 1, sqlite3_vec_sql_stmt_create_table(i), " ") - (Len("CREATE TABLE ") + 1))
    
    If sqlite3_check_if_table_already_exist(dbHandle, tmp_table_name) = True Then
        If drop_if_already_exists = True Then
            SQLite3PrepareV2 dbHandle, "DROP TABLE IF EXISTS " & tmp_table_name, stmHandle
            SQLite3Step stmHandle
            SQLite3Finalize stmHandle
            
            need_to_create_db = True
        Else
            need_to_create_db = False
        End If
    Else
        need_to_create_db = True
    End If
    
    If need_to_create_db = True Then
        SQLite3PrepareV2 dbHandle, sqlite3_vec_sql_stmt_create_table(i), stmHandle
        retValue = SQLite3Step(stmHandle)
        
        If retValue <> SQLITE_DONE Then
            sqlite3_create_tables = False
        End If
        
        SQLite3Finalize stmHandle
    End If
    
Next i


If IsNumeric(db_path) Then
    'laisse le pont ouvert
Else
    SQLite3Close dbHandle
End If

End Function


Function sqlite3_check_if_table_already_exist(ByVal db_path As Variant, ByVal table_name As String)

sqlite3_check_if_table_already_exist = False

Dim i As Double, j As Double

Dim initReturn As Long
Dim dbHandle As Long
Dim stmHandle As Long
Dim retValue As Long

If IsNumeric(db_path) Then
    dbHandle = db_path
Else
    initReturn = SQLite3Initialize(src_path & folder_dlls)
    retValue = SQLite3Open(db_path, dbHandle)
    retValue = SQLite3Finalize(stmHandle)
End If


Dim sql_query As String
sql_query = "SELECT name FROM sqlite_master"

Dim extract_master As Variant
extract_master = sqlite3_query(dbHandle, sql_query)

If IsEmpty(extract_master(0)) = True Then
    sqlite3_check_if_table_already_exist = False
Else
    For i = 0 To UBound(extract_master, 1)
        If extract_master(i)(0) = table_name Then
            sqlite3_check_if_table_already_exist = True
            Exit Function
        Else
            If i = UBound(extract_master, 1) Then
                sqlite3_check_if_table_already_exist = False
            End If
        End If
    Next i
End If

If IsNumeric(db_path) Then
    'laisse le pont ouvert
Else
    SQLite3Close dbHandle
End If

End Function



Function sqlite3_get_query_create_table(ByVal table_name As String, ByVal fields_and_options As Variant, Optional ByVal primary_keys As Variant) As Variant

Dim i As Double

'patch les types
For i = 0 To UBound(fields_and_options, 1)
    fields_and_options(i)(1) = sqlite3_get_field_type(fields_and_options(i)(1))
Next i


Dim sql_query As String
sql_query = ""

sql_query = "CREATE TABLE " & table_name & " ("

For i = 0 To UBound(fields_and_options, 1)
    
    sql_query = sql_query & fields_and_options(i)(0) & " " & fields_and_options(i)(1)
    
    If fields_and_options(i)(2) <> "" Then
        sql_query = sql_query & " " & fields_and_options(i)(2)
    End If
    
    If i <> UBound(fields_and_options, 1) Then
        sql_query = sql_query & ", "
    End If
    
Next i


Dim primary_key_str As String
primary_key_str = ""
If IsMissing(primary_keys) = False Then
    
    primary_key_str = ", PRIMARY KEY("
    
    For i = 0 To UBound(primary_keys, 1)
        primary_key_str = primary_key_str & primary_keys(i)(0) & " " & primary_keys(i)(1)
        
        If i <> UBound(primary_keys, 1) Then
            primary_key_str = primary_key_str & ", "
        End If
    Next i
    
    primary_key_str = primary_key_str & ")"
    
    sql_query = sql_query & " " & primary_key_str
End If


sql_query = sql_query & ")"


sqlite3_get_query_create_table = sql_query

End Function




Function sqlite3_query(ByVal db_path As Variant, ByVal query As String, Optional ByVal without_header As Boolean = False) As Variant

Dim i As Double, j As Double, k As Double, m As Double, n As Double

Dim src_path As String
src_path = StrReverse(Mid(StrReverse(ThisWorkbook.FullName), InStr(StrReverse(ThisWorkbook.FullName), "\")))

Dim initReturn As Long

Dim dbHandle As Long
Dim stmHandle As Long
Dim retValue As Long

If IsNumeric(db_path) Then
    dbHandle = db_path
Else
    initReturn = SQLite3Initialize(src_path & folder_dlls)
    retValue = SQLite3Open(db_path, dbHandle)
    retValue = SQLite3Finalize(stmHandle)
End If


retValue = SQLite3PrepareV2(dbHandle, query, stmHandle)
retValue = SQLite3Step(stmHandle)

Dim matrix_data() As Variant
ReDim Preserve matrix_data(0)

Dim first_row As Boolean
first_row = True

Dim colCount As Integer
colCount = SQLite3ColumnCount(stmHandle)

Dim vec_tmp_data() As Variant

While (retValue = SQLITE_ROW)

    If first_row = True And without_header = False Then
        
        ReDim vec_tmp_data(colCount - 1)
        For i = 0 To colCount - 1
            vec_tmp_data(i) = SQLite3ColumnName(stmHandle, i)
        Next i
        
        matrix_data(0) = vec_tmp_data
        
        
        
        If first_row = True Then
            k = 1
        End If
        
        first_row = False
    Else
        If first_row = True Then
            k = 0
        End If
        
        first_row = False
    End If
    
    For i = 0 To colCount - 1
        ReDim Preserve vec_tmp_data(i)
        vec_tmp_data(i) = ColumnValue(stmHandle, i, SQLite3ColumnType(stmHandle, i))
    Next i
    
    ReDim Preserve matrix_data(k)
    matrix_data(k) = vec_tmp_data
    k = k + 1
    
    retValue = SQLite3Step(stmHandle)
Wend


'fermeture du point
If IsNumeric(db_path) Then
    
Else
    retValue = SQLite3Close(dbHandle)
End If

sqlite3_query = matrix_data

End Function


'les fonctions peuvent recevoir le path de la DB ou alors un dbHandle
Function sqlite3_query_vec(ByVal db_path As Variant, ByVal vec_query_create_table As Variant, Optional ByVal drop_if_exists As Boolean = False) As Boolean

Dim i As Double, j As Double, k As Double, m As Double, n As Double

Dim src_path As String
src_path = StrReverse(Mid(StrReverse(ThisWorkbook.FullName), InStr(StrReverse(ThisWorkbook.FullName), "\")))

Dim initReturn As Long
initReturn = SQLite3Initialize(src_path & folder_dlls)

'ouverture du pont
Dim dbHandle As Long
Dim stmHandle As Long
Dim retValue As Long

If IsNumeric(db_path) Then
    dbHandle = db_path
Else
    retValue = SQLite3Open(db_path, dbHandle)
    retValue = SQLite3Finalize(stmHandle)
End If



End Function


Function sqlite3_adjust_structure(ByVal db_path As Variant, ByVal table_name As String, ByVal field_name_and_type As Variant) As Variant

Dim i As Double, j As Double, k As Double, m As Double, n As Double

Dim Status As Boolean

'patch les type
For i = 0 To UBound(field_name_and_type, 1)
    field_name_and_type(i)(1) = sqlite3_get_field_type(field_name_and_type(i)(1))
Next i


Dim src_path As String
src_path = StrReverse(Mid(StrReverse(ThisWorkbook.FullName), InStr(StrReverse(ThisWorkbook.FullName), "\")))

Dim initReturn As Long

Dim dbHandle As Long
Dim stmHandle As Long
Dim retValue As Long

If IsNumeric(db_path) Then
    dbHandle = db_path
Else
    initReturn = SQLite3Initialize(src_path & folder_dlls)
    retValue = SQLite3Open(db_path, dbHandle)
    retValue = SQLite3Finalize(stmHandle)
End If


Dim sql_query As String


'mount la stucture actuelle

Dim actual_structure As Variant
extract_structure = sqlite3_get_table_structure(db_path, table_name)

'transforme en un simple vecteur les differents champs
Dim vec_field() As Variant
k = 0

If IsEmpty(extract_structure(0)) = False Then
    
    Dim dim_structure_name As Integer
    
    
    'repere la dim name
    For i = 0 To UBound(extract_structure(0), 1)
        If extract_structure(0)(i) = "name" Then
            dim_structure_name = i
            Exit For
        End If
    Next i
    
    For i = 1 To UBound(extract_structure, 1)
        ReDim Preserve vec_field(i - 1)
        vec_field(i - 1) = extract_structure(i)(dim_structure_name)
    Next i
    
    
Else
    sqlite3_adjust_structure = -1
    Exit Function
End If


'fusion entre need_field et vec_field
Dim found_field As Boolean
k = 0
For i = 0 To UBound(field_name_and_type, 1)
    found_field = False
    For j = 0 To UBound(vec_field, 1)
        If field_name_and_type(i)(0) = vec_field(j) Then
            found_field = True
            Exit For
        End If
    Next j
    
    Dim field_to_add_to_table() As Variant
    
    
    If found_field = False Then
        ReDim Preserve field_to_add_to_table(k)
        field_to_add_to_table(k) = field_name_and_type(i)
        k = k + 1
    End If
    
Next i

If k > 0 Then
    'mise a jour de la structure de la table
    Status = sqlite3_insert_column(dbHandle, table_name, field_to_add_to_table)
    
    sqlite3_adjust_structure = field_to_add_to_table
Else
    sqlite3_adjust_structure = 0
End If


If IsNumeric(db_path) Then
    'laisse le pont ouvert
Else
    SQLite3Close dbHandle
End If


End Function


Function sqlite3_insert_column(ByVal db_path As Variant, ByVal table_name As String, ByVal field_name_and_type As Variant) As Boolean

Dim i As Double, j As Double, k As Double, m As Double, n As Double


'patch les type
For i = 0 To UBound(field_name_and_type, 1)
    field_name_and_type(i)(1) = sqlite3_get_field_type(field_name_and_type(i)(1))
Next i


Dim src_path As String
src_path = StrReverse(Mid(StrReverse(ThisWorkbook.FullName), InStr(StrReverse(ThisWorkbook.FullName), "\")))

Dim initReturn As Long

Dim dbHandle As Long
Dim stmHandle As Long
Dim retValue As Long

If IsNumeric(db_path) Then
    dbHandle = db_path
Else
    initReturn = SQLite3Initialize(src_path & folder_dlls)
    retValue = SQLite3Open(db_path, dbHandle)
    retValue = SQLite3Finalize(stmHandle)
End If


Dim sql_query As String
retValue = SQLite3PrepareV2(dbHandle, "BEGIN TRANSACTION", stmHandle)
retValue = SQLite3Step(stmHandle)
retValue = SQLite3Finalize(stmHandle)

For i = 0 To UBound(field_name_and_type, 1)
    sql_query = "ALTER TABLE " & table_name & " ADD " & field_name_and_type(i)(0) & " " & field_name_and_type(i)(1)
    
    retValue = SQLite3PrepareV2(dbHandle, sql_query, stmHandle)
    retValue = SQLite3Step(stmHandle)
    retValue = SQLite3Reset(stmHandle)
Next i



retValue = SQLite3Finalize(stmHandle)

retValue = SQLite3PrepareV2(dbHandle, "COMMIT TRANSACTION", stmHandle)
retValue = SQLite3Step(stmHandle)
sqlite3_insert_column = retValue
SQLite3Finalize stmHandle

If IsNumeric(db_path) Then
    'laisse le pont ouvert
Else
    SQLite3Close dbHandle
End If

End Function




Function sqlite3_get_table_structure(ByVal db_path As Variant, ByVal table_name As String) As Variant

Dim i As Double, j As Double, k As Double, m As Double, n As Double

Dim src_path As String
src_path = StrReverse(Mid(StrReverse(ThisWorkbook.FullName), InStr(StrReverse(ThisWorkbook.FullName), "\")))

Dim initReturn As Long

Dim dbHandle As Long
Dim stmHandle As Long
Dim retValue As Long

If IsNumeric(db_path) Then
    dbHandle = db_path
Else
    initReturn = SQLite3Initialize(src_path & folder_dlls)
    retValue = SQLite3Open(db_path, dbHandle)
    retValue = SQLite3Finalize(stmHandle)
End If


Dim sql_query As String
sql_query = "SELECT * FROM sql_master"
sql_query = "PRAGMA table_info(""" & table_name & """)"

Dim extract_structure As Variant
extract_structure = sqlite3_query(dbHandle, sql_query)

'transforme en un simple vecteur les differents champs
Dim vec_field() As Variant
k = 0
If IsEmpty(extract_structure(0)) = False Then
    
    Dim dim_structure_name As Integer
    
    
    'repere la dim name
    For i = 0 To UBound(extract_structure(0), 1)
        If extract_structure(0)(i) = "name" Then
            dim_structure_name = i
            Exit For
        End If
    Next i
    
    If UBound(extract_structure, 1) > 0 Then
        For i = 1 To UBound(extract_structure, 1)
            ReDim Preserve vec_field(k)
            vec_field(k) = extract_structure(i)(dim_structure_name)
            k = k + 1
        Next i
        
        sqlite3_get_table_structure = vec_field
        sqlite3_get_table_structure = extract_structure
        'Exit Function
        
    End If
Else
    sqlite3_get_table_structure = -1
    'Exit Function
End If

If IsNumeric(db_path) Then
    'laisse le pont ouvert
Else
    SQLite3Close dbHandle
End If

End Function


Function sqlite3_insert_with_transaction(ByVal db_path As Variant, ByVal table_name As String, ByVal matrix_data As Variant, Optional ByVal field_name As Variant, Optional ByVal field_type As Variant) As Long

Dim i As Double, j As Double, k As Double, m As Double, n As Double

If IsArray(matrix_data(0)) Then
    Dim tmp_matrix As Variant
    tmp_matrix = vec_to_array(matrix_data)
    matrix_data = tmp_matrix
End If

Dim src_path As String
src_path = StrReverse(Mid(StrReverse(ThisWorkbook.FullName), InStr(StrReverse(ThisWorkbook.FullName), "\")))

Dim initReturn As Long

Dim dbHandle As Long
Dim stmHandle As Long
Dim retValue As Long

If IsNumeric(db_path) Then
    dbHandle = db_path
Else
    initReturn = SQLite3Initialize(src_path & folder_dlls)
    retValue = SQLite3Open(db_path, dbHandle)
    retValue = SQLite3Finalize(stmHandle)
End If


Dim sql_query As String

Dim extract_structure As Variant
Dim dim_structure_name As Integer, dim_structure_type As Integer

If IsMissing(field_name) Then
    sql_query = "PRAGMA table_info(""" & table_name & """)"
    
    extract_structure = sqlite3_query(dbHandle, sql_query)
    
    For i = 0 To UBound(extract_structure(0), 1)
        If extract_structure(0)(i) = "type" Then
            dim_structure_type = i
        ElseIf extract_structure(0)(i) = "name" Then
            dim_structure_name = i
        End If
    Next i
    
    ReDim field_name(UBound(extract_structure, 1) - 1)
    For i = 1 To UBound(extract_structure, 1)
        field_name(i - 1) = extract_structure(i)(dim_structure_name)
    Next i
    
End If



If IsMissing(field_type) Then
    
    ReDim field_type(UBound(field_name, 1))
    
    'extraction directement de la structure de la table
    sql_query = "PRAGMA table_info(""" & table_name & """)"
    
    extract_structure = sqlite3_query(dbHandle, sql_query)
    
    For i = 0 To UBound(extract_structure(0), 1)
        If extract_structure(0)(i) = "type" Then
            dim_structure_type = i
        ElseIf extract_structure(0)(i) = "name" Then
            dim_structure_name = i
        End If
    Next i
    
    'pour chaque champs repere son type
    For i = 0 To UBound(field_name, 1)
        For j = 1 To UBound(extract_structure, 1)
            If field_name(i) = extract_structure(j)(dim_structure_name) Then
                field_type(i) = extract_structure(j)(dim_structure_type)
                Exit For
            End If
        Next j
    Next i
    
End If


SQLite3PrepareV2 dbHandle, "BEGIN TRANSACTION", stmHandle
SQLite3Step stmHandle
SQLite3Finalize stmHandle



sql_query = "INSERT INTO " & table_name & " ("

Dim insert_point As String
insert_point = ""
For i = 0 To UBound(field_name, 1)
    sql_query = sql_query & field_name(i)
    insert_point = insert_point & "?"
    
    If i <> UBound(field_name, 1) Then
        sql_query = sql_query & ", "
        insert_point = insert_point & ", "
    End If
Next i

sql_query = sql_query & ") VALUES (" & insert_point & ")"

SQLite3PrepareV2 dbHandle, sql_query, stmHandle


For i = 0 To UBound(matrix_data, 1)
    For j = 0 To UBound(matrix_data, 2)
        
        If IsNull(field_type(j)) Or IsEmpty(field_type(j)) Or IsEmpty(matrix_data(i, j)) Or IsNull(matrix_data(i, j)) Then
            SQLite3BindNull stmHandle, j + 1
        ElseIf UCase(field_type(j)) = "INT" Or UCase(field_type(j)) = "INTEGER" Then
            SQLite3BindInt32 stmHandle, j + 1, CInt(matrix_data(i, j))
        ElseIf UCase(field_type(j)) = "DOUBLE" Or UCase(field_type(j)) = "REAL" Then
            SQLite3BindDouble stmHandle, j + 1, CDbl(matrix_data(i, j))
        ElseIf UCase(field_type(j)) = "DATE" Then
            SQLite3BindDate stmHandle, j + 1, matrix_data(i, j)
        ElseIf UCase(field_type(j)) = "NUMERIC" Then
            If IsDate(matrix_data(i, j)) Then
                SQLite3BindDate stmHandle, j + 1, matrix_data(i, j)
            Else
                SQLite3BindDouble stmHandle, j + 1, CDbl(matrix_data(i, j))
            End If
        ElseIf UCase(field_type(j)) = "TEXT" Then
            retValue = SQLite3BindText(stmHandle, j + 1, CStr(matrix_data(i, j)))
        ElseIf UCase(field_type(j)) = "BOOL" Then
            If matrix_data(i, j) = False Or matrix_data(i, j) = 0 Then
                SQLite3BindInt32 stmHandle, j + 1, 0
            Else
                SQLite3BindInt32 stmHandle, j + 1, 1
            End If
        End If
        
    Next j
    
    retValue = SQLite3Step(stmHandle)
    retValue = SQLite3Reset(stmHandle)
    
Next i

SQLite3Finalize stmHandle

SQLite3PrepareV2 dbHandle, "COMMIT TRANSACTION", stmHandle
retValue = SQLite3Step(stmHandle)
sqlite3_insert_with_transaction = retValue
SQLite3Finalize stmHandle


If IsNumeric(db_path) Then
    'laisse le pont ouvert
Else
    SQLite3Close dbHandle
End If

End Function

Public Function sqlite3_str_to_date(ByVal dt_str As String) As Date

sqlite3_str_to_date = Right(dt_str, 2) & "." & Mid(dt_str, 6, 2) & "." & Left(dt_str, 4)

End Function


Public Function sqlite3_date_to_str(ByVal dt_date As Date) As String

sqlite3_date_to_str = Right(dt_date, 4) & "-" & Mid(dt_date, 4, 2) & "-" & Left(dt_date, 2)

End Function


Public Function vec_to_array(ByVal vec As Variant) As Variant

Dim i As Double, j As Double
Dim matrix_data() As Variant


'repere la plus grande dim
Dim max_dim As Integer
max_dim = 0
For i = 0 To UBound(vec, 1)
    If UBound(vec(i), 1) > max_dim Then
        max_dim = UBound(vec(i), 1)
    End If
Next i

ReDim matrix_data(UBound(vec, 1), max_dim)

For i = 0 To UBound(vec, 1)
    For j = 0 To UBound(vec(i), 1)
        matrix_data(i, j) = vec(i)(j)
    Next j
Next i

vec_to_array = matrix_data

End Function


Public Function array_to_vec(ByVal matrix As Variant) As Variant

Dim i As Double, j As Double
Dim vec() As Variant
ReDim vec(UBound(matrix, 1))

Dim tmp_sub_vec() As Variant
For i = 0 To UBound(matrix, 1)
    ReDim tmp_sub_vec(UBound(matrix, 2))
    
    For j = 0 To UBound(matrix, 2)
        tmp_sub_vec(j) = matrix(i, j)
    Next j
    
    vec(i) = tmp_sub_vec
Next i

array_to_vec = vec

End Function

