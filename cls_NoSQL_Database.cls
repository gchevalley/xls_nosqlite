VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_NoSQL_Database"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public local_filepath As String 'equivalent du name ?

Public sqlite_prefix_table_name As String
Public sqlite_field_oid As String
Public sqlite_field_json_data As String

Public oOId As cls_NoSQL_ObjectId
    'Private last_timestamp As Long
    'Private objectid_count As Long
    
Public collections As Scripting.Dictionary

'Private Const nosql_encode_char_json As String = "*#"
Private Const nosql_encode_char_json As String = "'"


Private Sub Class_Initialize()

Me.local_filepath = ""

' #################### DON'T TOUCH ####################
Me.sqlite_prefix_table_name = "nosql_collection_"
Me.sqlite_field_oid = "_id"
Me.sqlite_field_json_data = "json_data"
' #####################################################


Set Me.oOId = New cls_NoSQL_ObjectId
    'last_timestamp = Me.oOId.get_current_timestamp_unix()
    'objectid_count = 0


Set collections = New Scripting.Dictionary 'pour la partie purement excel -> doit etre sync tout le temps avec DB

End Sub


Public Function setup_with_file(ByVal filepath As String)

local_filepath = filepath
'sqlite3_create_db filepath
create_new_database

'mount les tables


End Function


Private Function create_new_database()

sqlite3_create_db Me.local_filepath

'creation du helper
Dim sql_query_create_table As String
If sqlite3_check_if_table_already_exist(Me.local_filepath, "helper") = False Then
    
    Debug.Print "INFO cls_NoSQL_Database::create_new_database because " & Me.local_filepath & " is a new database, need to create helper table"
    
    sql_query_create_table = sqlite3_get_query_create_table("helper", Array(Array(Me.sqlite_field_oid, "TEXT", ""), Array(Me.sqlite_field_json_data, "TEXT", "")), Array(Array(Me.sqlite_field_oid, "ASC")))
    
    sqlite3_create_tables Me.local_filepath, Array(sql_query_create_table)
End If

End Function


Public Function get_sqlite_table_name_equivalent_from_collection_name(ByVal collection_name As String)

get_sqlite_table_name_equivalent_from_collection_name = Me.sqlite_prefix_table_name & collection_name

End Function


Public Function check_collection_validity_name(ByVal collection_name As String) As Boolean

Dim i As Integer

check_collection_validity_name = True

'check name validity
Dim forbidden_char_collection_name() As Variant
    forbidden_char_collection_name = Array(" ", "/", "\", "#", "@", "&", "+", "(", ")", "=", """", "?", "!", "-", "*", "%")

For i = 0 To UBound(forbidden_char_collection_name, 1)
    If InStr(UCase(collection_name), UCase(forbidden_char_collection_name(i))) <> 0 Then
        check_collection_validity_name = False
        Exit Function
    End If
Next i

End Function


Public Function use(ByVal collection_name As String) As cls_NoSQL_Collection

Debug.Print "INFO enter cls_NoSQL_Database::use with collection_name: " & collection_name

Dim i As Long, j As Long, k As Long

If check_collection_validity_name(collection_name) = False Then
    Debug.Print "ERROR cls_NoSQL_Database::use " & "invalid name for collection name: " & collection_name
    MsgBox ("invalid name for collection name: " & collection_name)
    Exit Function
End If


'deja presente ?
If sqlite3_check_if_table_already_exist(Me.local_filepath, get_sqlite_table_name_equivalent_from_collection_name(collection_name)) = True Then
    
    Debug.Print "INFO cls_NoSQL_Database::use collection_name: " & collection_name & " already existing as sqlite table"
    
    Set use = New cls_NoSQL_Collection
    use.name = collection_name
    use.link_with_database Me
    
    
    'avec des donnees ?
    Dim oJSON As New jsonlib
    Dim extract_table_content As Variant
    extract_table_content = sqlite3_query(Me.local_filepath, "SELECT _id, json_data FROM " & get_sqlite_table_name_equivalent_from_collection_name(collection_name))
    
    If UBound(extract_table_content, 1) > 0 Then
        
        Debug.Print "INFO cls_NoSQL_Database::use collection_name: " & collection_name & " contains " & UBound(extract_table_content, 1) & " documents based on row of sqlite table"
        
        For i = 1 To UBound(extract_table_content, 1)
            
            If InStr(extract_table_content(i)(0), "lg") <> 0 Or InStr(extract_table_content(i)(0), "mobile2") <> 0 Then
                Debug.Print extract_table_content(i)(0)
            End If
            
            use.load_document nosql_decode_json_from_DB(extract_table_content(i)(1)), nosql_decode_json_from_DB(extract_table_content(i)(0))
        Next i
        
    Else
        
        Debug.Print "INFO cls_NoSQL_Database::use collection_name: " & collection_name & " is currently empty"
        
    End If
    
Else 'sinon la cree
    
    Debug.Print "INFO cls_NoSQL_Database::use need to create a new collection_name: " & collection_name
    
    Set use = create_new_collection(collection_name)
    
End If



End Function


Private Function create_new_collection(ByVal collection_name As String) As cls_NoSQL_Collection

Debug.Print "INFO enter cls_NoSQL_Database::create_new_collection with collection_name: " & collection_name

'prepare la table d accueil sqlite
Dim sqlite_query As String
sqlite_query = get_sqlite_create_table_query_for_new_collection(collection_name)
sqlite3_create_tables Me.local_filepath, Array(get_sqlite_create_table_query_for_new_collection(collection_name))

Set create_new_collection = New cls_NoSQL_Collection
create_new_collection.name = collection_name
create_new_collection.link_with_database Me

End Function


Private Function get_sqlite_create_table_query_for_new_collection(ByVal collection_name As String)

get_sqlite_create_table_query_for_new_collection = sqlite3_get_query_create_table(get_sqlite_table_name_equivalent_from_collection_name(collection_name), Array(Array(Me.sqlite_field_oid, "TEXT", ""), Array(Me.sqlite_field_json_data, "TEXT", "")), Array(Array(Me.sqlite_field_oid, "ASC")))

End Function



Public Function nosql_encode_json_for_DB(ByVal str As String) As String

nosql_encode_json_for_DB = Replace(str, """", nosql_encode_char_json)

End Function


Private Function nosql_decode_json_from_DB(ByVal str As String) As String

nosql_decode_json_from_DB = Replace(str, nosql_encode_char_json, """")

End Function




Private Sub Class_Terminate()

Set Me.collections = Nothing

End Sub
