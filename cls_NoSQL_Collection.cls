VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_NoSQL_Collection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public name As String
Public documents As Scripting.Dictionary
Public ref_database As cls_NoSQL_Database


Private Sub Class_Initialize()

Set Me.documents = New Scripting.Dictionary

Me.name = ""
Set Me.ref_database = New cls_NoSQL_Database

End Sub


Public Function link_with_database(oDB As cls_NoSQL_Database)

Debug.Print "INFO cls_NoSQL_Collection::link_with_database collection: " & Me.name & " with database: " & oDB.local_filepath

Set Me.ref_database = oDB

End Function


'first level should absolutely an object
Public Function insert(o) 'json string or excel's dictionary

Dim oJSON As New jsonlib

Dim tmp_document As cls_NoSQL_Document
    Set tmp_document = New cls_NoSQL_Document

tmp_document.load_data o
    

'append object id
If tmp_document.representation_dictionary.Exists("_id") Then
    Debug.Print "WARNING cls_NoSQL_Collection::insert document: " & tmp_document.representation_json & " already contains an _id, need to remove it"
    tmp_document.representation_dictionary.remove "_id"
End If


'Debug.Print Me.ref_database.oOId.get_next()

tmp_document.representation_dictionary.Add "_id", Me.ref_database.oOId.get_next_dic()


'regenere le doc pour assurer sync entre les 2 represenation
Dim final_document_to_store As cls_NoSQL_Document
    Set final_document_to_store = New cls_NoSQL_Document
final_document_to_store.load_data tmp_document.representation_dictionary


'Debug.Print final_document_to_store.representation_json
'Dim check_oid_as_dic As Scripting.Dictionary
'Set check_oid_as_dic = final_document_to_store.representation_dictionary.Item("_id")
Set tmp_document = Nothing

'sync avec DB sqlite, la chaine json a besoin au prealable d un traitement pour pouvoir etre stocker dans la table sqlite
sqlite3_query Me.ref_database.local_filepath, "INSERT INTO " & Me.ref_database.get_sqlite_table_name_equivalent_from_collection_name(Me.name) & " (" & Me.ref_database.sqlite_field_oid & ", " & Me.ref_database.sqlite_field_json_data & ") VALUES (""" & Me.ref_database.nosql_encode_json_for_DB(oJSON.toString(final_document_to_store.representation_dictionary.Item("_id"))) & """,""" & Me.ref_database.nosql_encode_json_for_DB(final_document_to_store.representation_json) & """)"

'mise a jour de l object excel
Me.load_document final_document_to_store

End Function


Public Function find(oQueryJSON, Optional limit_nbre_items As Variant, Optional sort As Variant, Optional select_fields As Variant) As cls_NoSQL_QueryResult

Set find = New cls_NoSQL_QueryResult
    Set find.ref_collection = Me

Dim i As Long, j As Long, k As Long

Dim oJSON As New jsonlib

Dim oReg As New VBScript_RegExp_55.RegExp
Dim match As VBScript_RegExp_55.match
Dim matches As VBScript_RegExp_55.MatchCollection
    oReg.Global = True

Dim final_query_json As Scripting.Dictionary
Set final_query_json = transform_query_into_object(oQueryJSON)

If final_query_json Is Nothing Then
    Debug.Print "ERROR cls_NoSQL_Collection::find with oQueryJSON -> Quit"
    Exit Function
Else
    find.query = oJSON.toString(final_query_json)
End If


'traitement sur select_fields
Dim checked_select_fields As Boolean
    checked_select_fields = False
Dim final_select_fields As Scripting.Dictionary
If IsMissing(select_fields) = False Then
    Set final_select_fields = transform_query_into_object(select_fields)
    If final_select_fields Is Nothing Then
        checked_select_fields = False
    End If
End If


'traitement sur l array de sort
Dim final_sort_fields As Collection 'attention l ordre compte -> ne doit pas etre un dico mais une collection
If IsMissing(sort) = False Then
    'set final_sort_fields
End If


'passe en revue les documents
Dim tmp_oid As Variant
Dim oDoc As cls_NoSQL_Document
Dim take_item As Boolean
Dim find_candidate As Boolean

Dim dic_constraints_find_candidate As New Scripting.Dictionary

For Each tmp_oid In Me.documents.keys
    
    Set oDoc = Me.documents.Item(tmp_oid)
    
    'passe en revue les contraintes
    Set dic_constraints_find_candidate = New Scripting.Dictionary
    Dim tmp_constraint_field As Variant
    For Each tmp_constraint_field In final_query_json.keys
        
        Dim col_objects_to_check As Collection 'un obet candidat contient juste tout le schema de cles necessaire et se trouve au dernier niveau il ne reste qu a checker la contrainte
        Set col_objects_to_check = oDoc.get_collection_sub_objects(tmp_constraint_field)
        
        'Dim col_final_candidates As Collection non implementer pour gagner en rapidite d execution sur les recherches
        
        Dim dic_CheckLevelConstraint As Scripting.Dictionary ', dic_current_level As Scripting.Dictionary, dic_sub_level As Scripting.Dictionary
        
        find_candidate = False
        dic_constraints_find_candidate.Add tmp_constraint_field, False
        
        If col_objects_to_check.count > 0 Then
            
            Dim tmp_constraint_field_last_level As String
            tmp_constraint_field_last_level = find_get_last_level_on_query_field(tmp_constraint_field)
            
            For Each dic_CheckLevelConstraint In col_objects_to_check
                
                If find_candidate = False Then
        
                    If dic_CheckLevelConstraint.Exists(tmp_constraint_field_last_level) Then
                        
                        'type de contrainte scalaire, complexe -> egalement checker operator, optimiser avec regexp
                        If (VarType(final_query_json.Item(tmp_constraint_field)) = vbObject And TypeOf final_query_json.Item(tmp_constraint_field) Is Scripting.Dictionary) Then
                            
                            Dim tmp_sub_constraint As Scripting.Dictionary
                            
                            'transforme en dico si necessaire
                            If (VarType(final_query_json.Item(tmp_constraint_field)) = vbObject And TypeOf final_query_json.Item(tmp_constraint_field) Is Scripting.Dictionary) Then
                                Set tmp_sub_constraint = final_query_json.Item(tmp_constraint_field)
                            Else
                                Set tmp_sub_constraint = oJSON.parse(CStr(final_query_json.Item(tmp_constraint_field)))
                            End If
                            
                            
                            'repere l operator
                            Dim tmp_key_sub_constraint As Variant 'devoir y en avoir qu une seule
                            For Each tmp_key_sub_constraint In tmp_sub_constraint.keys
                                
                                Select Case tmp_key_sub_constraint
                                
                                    Case "$gt"
                                        If IsNumeric(tmp_sub_constraint.Item(tmp_key_sub_constraint)) Then
                                            If IsNumeric(dic_CheckLevelConstraint.Item(tmp_constraint_field_last_level)) Then
                                                If dic_CheckLevelConstraint.Item(tmp_constraint_field_last_level) > tmp_sub_constraint.Item(tmp_key_sub_constraint) Then
                                                    find_candidate = True
                                                End If
                                            End If
                                        Else
                                            If dic_CheckLevelConstraint.Item(tmp_constraint_field_last_level) > tmp_sub_constraint.Item(tmp_key_sub_constraint) Then
                                                find_candidate = True
                                            End If
                                        End If
                                    
                                    
                                    Case "$gte"
                                        If IsNumeric(tmp_sub_constraint.Item(tmp_key_sub_constraint)) Then
                                            If IsNumeric(dic_CheckLevelConstraint.Item(tmp_constraint_field_last_level)) Then
                                                If dic_CheckLevelConstraint.Item(tmp_constraint_field_last_level) >= tmp_sub_constraint.Item(tmp_key_sub_constraint) Then
                                                    find_candidate = True
                                                End If
                                            End If
                                        Else
                                            If dic_CheckLevelConstraint.Item(tmp_constraint_field_last_level) >= tmp_sub_constraint.Item(tmp_key_sub_constraint) Then
                                                find_candidate = True
                                            End If
                                        End If
                                    
                                    
                                    Case "$lt"
                                        If IsNumeric(tmp_sub_constraint.Item(tmp_key_sub_constraint)) Then
                                            If IsNumeric(dic_CheckLevelConstraint.Item(tmp_constraint_field_last_level)) Then
                                                If dic_CheckLevelConstraint.Item(tmp_constraint_field_last_level) < tmp_sub_constraint.Item(tmp_key_sub_constraint) Then
                                                    find_candidate = True
                                                End If
                                            End If
                                        Else
                                            If dic_CheckLevelConstraint.Item(tmp_constraint_field_last_level) < tmp_sub_constraint.Item(tmp_key_sub_constraint) Then
                                                find_candidate = True
                                            End If
                                        End If
                                    
                                    
                                    Case "$lte"
                                        If IsNumeric(tmp_sub_constraint.Item(tmp_key_sub_constraint)) Then
                                            If IsNumeric(dic_CheckLevelConstraint.Item(tmp_constraint_field_last_level)) Then
                                                If dic_CheckLevelConstraint.Item(tmp_constraint_field_last_level) <= tmp_sub_constraint.Item(tmp_key_sub_constraint) Then
                                                    find_candidate = True
                                                End If
                                            End If
                                        Else
                                            If dic_CheckLevelConstraint.Item(tmp_constraint_field_last_level) <= tmp_sub_constraint.Item(tmp_key_sub_constraint) Then
                                                find_candidate = True
                                            End If
                                        End If
                                    
                                    
                                    'Case "$all"
                                    
                                    
                                    Case Else
                                        Debug.Print "WARNING cls_NoSQL_Collection::find unkown operator " & tmp_key_sub_constraint
                                End Select
                            Next
                        
                        Else
                            ' simple check equality
                            If dic_CheckLevelConstraint.Item(tmp_constraint_field_last_level) = final_query_json.Item(tmp_constraint_field) Then
                                find_candidate = True
                            End If
                            
                        End If
                        
                        
                    End If
                
                End If
            
            Next 'loop candidate dans (sous-) objet ?
            
            
            If find_candidate = True Then ' y a t il au moins un (sous-) canditat ?
                dic_constraints_find_candidate.Item(tmp_constraint_field) = True
            End If
            
        Else
            
            take_item = False ' aucun objet dans la collection
        
        End If 'y a t il au moins un sous-objet ?
        
    Next 'loop constraint
    
    
    'toutes les contraintes sont elles respectee ?
    Dim tmp_constraints As Variant
    take_item = True
    For Each tmp_constraint_field In dic_constraints_find_candidate.keys
        If take_item = True And dic_constraints_find_candidate.Item(tmp_constraint_field) = False Then
            take_item = False
            Exit For
        End If
    Next
    
    
check_state_take_item:
    If take_item Then
        
        If IsMissing(select_fields) = True Then
add_full_doc_to_result:
    
            'Debug.Print oDoc.representation_json
            
            find.append_document oDoc
        Else
            
            If checked_select_fields = True Then
            
                Dim tmp_dic_with_filtered_fields As Scripting.Dictionary
                    Set tmp_dic_with_filtered_fields = New Scripting.Dictionary
                        
                        'Debug.Print oDoc.representation_dictionary.Item("_id")
                        
                    tmp_dic_with_filtered_fields.Add "_id", oJSON.parse(CStr(oDoc.representation_dictionary.Item("_id"))) 'en tout cas l id
                    
                'passage en revue des champs souhaite par l utilisateur
                Dim tmp_desired_field As Variant
                For Each tmp_desired_field In final_select_fields.keys
                    
                    If IsNumeric(final_select_fields.Item(tmp_desired_field)) Then
                        If final_select_fields.Item(tmp_desired_field) = 1 Then
                            
                            If oDoc.representation_dictionary.Exists(tmp_desired_field) Then
                                tmp_dic_with_filtered_fields.Add tmp_desired_field, oDoc.representation_dictionary.Item(tmp_desired_field)
                            End If
                        End If
                    End If
                    
                Next
                
                Dim oDocWithSelectedField As cls_NoSQL_Document
                Set oDocWithSelectedField = New cls_NoSQL_Document
                    oDocWithSelectedField.link_with_collection Me
                    oDocWithSelectedField.load_data tmp_dic_with_filtered_fields
                
                find.append_document oDocWithSelectedField
                
            Else
                GoTo add_full_doc_to_result
            End If
            
        End If
    End If
check_next_document:
Next 'loop documents

Debug.Print "INFO cls_NoSQL_Collection::find query: " & oJSON.toString(final_query_json) & " returns: " & find.count & " result(s)"
 
End Function


Private Function find_get_last_level_on_query_field(ByVal query_field As String) As String

If InStr(query_field, ".") <> 0 Then
    find_get_last_level_on_query_field = StrReverse(Left(StrReverse(query_field), InStr(StrReverse(query_field), ".") - 1))
Else
    find_get_last_level_on_query_field = query_field
End If

End Function


' deprecie, directement dans objet document
'Private Function find_get_sub_object_candidates_based_on_query_field(oDoc As cls_NoSQL_Document, ByVal query_field As String) As Collection
'
'Set find_get_sub_object_candidates_based_on_query_field = New Collection
'
'Set find_get_sub_object_candidates_based_on_query_field = oDoc.get_collection_sub_objects(query_field)
'
''Dim k As Long
''
''Dim oReg As New VBScript_RegExp_55.RegExp
''Dim match As VBScript_RegExp_55.match
''Dim matches As VBScript_RegExp_55.MatchCollection
''    oReg.Global = True
''
''
'''embedded field ?
''Dim array_CheckLevelConstraint As New Collection, array_CheckSubLevelConstraint As Collection
''Dim dic_CheckLevelConstraint As Scripting.Dictionary ', dic_current_level As Scripting.Dictionary, dic_sub_level As Scripting.Dictionary
''
''Set dic_CheckLevelConstraint = oDoc.representation_dictionary 'top level must be a dictionary
''
''array_CheckLevelConstraint.Add dic_CheckLevelConstraint 'package tout ca dans une collection
''
''
''If InStr(query_field, ".") <> 0 Then
''
''    oReg.Pattern = "[\w]+"
''    Set matches = oReg.Execute(query_field)
''
''    k = 1
''    For Each match In matches
''
''        Set array_CheckSubLevelConstraint = New Collection
''
''        'passe en revue les objects (dictionary) de la collection
''        Dim tmp_o As Scripting.Dictionary
''
''        Dim take_candidate As Boolean
''        For Each tmp_o In array_CheckLevelConstraint
''
''            take_candidate = True
''
''
''            If tmp_o.Exists(match.value) Then 'ce level contient il la cle du sublevel suivant ?
''
''            Else
''                take_candidate = False
''            End If
''
''
''
''            If take_candidate = True Then
''
''                If VarType(tmp_o.Item(match.value)) = vbObject Then
''                    If TypeOf tmp_o.Item(match.value) Is Scripting.Dictionary Then
''                        array_CheckSubLevelConstraint.Add tmp_o.Item(match.value) 'single object
''                    ElseIf TypeOf tmp_o.Item(match.value) Is Collection Then
''                        Set array_CheckSubLevelConstraint = tmp_o.Item(match.value) 'array d objects
''                    End If
''                Else
''                    array_CheckSubLevelConstraint.Add tmp_o 'ne descend pas plus bas 'scalaire
''                End If
''
''            End If
''
''        Next
''
''        Set array_CheckLevelConstraint = array_CheckSubLevelConstraint 'next -> current
''
''    Next
''
''    Set find_get_sub_object_candidates_based_on_query_field = array_CheckLevelConstraint
''
''Else 'normal condition top level dictionary
''    find_get_sub_object_candidates_based_on_query_field.Add oDoc.representation_dictionary
''End If
'
'
'End Function


Public Function update(oSelectQuery, oUpdateQueryORFullNewDocument, Optional ByVal multi As Boolean = False) As cls_NoSQL_QueryResult

Set update = New cls_NoSQL_QueryResult

Dim oJSON As New jsonlib

Dim oReg As New VBScript_RegExp_55.RegExp
Dim match As VBScript_RegExp_55.match
Dim matches As VBScript_RegExp_55.MatchCollection
    oReg.Global = True


Dim oResultsSelect As cls_NoSQL_QueryResult
    Set oResultsSelect = Me.find(oSelectQuery) 'rapatrie les documents qui correspondent a la query de selection

Dim oDicUpdateQuery As Scripting.Dictionary
Set oDicUpdateQuery = transform_query_into_object(oUpdateQueryORFullNewDocument)


'distinct full new doc (erase old replace with new) or update modifier (just part(s) modifier)
Dim tmp_oid As Variant
If oDicUpdateQuery.count = 1 And Left(oDicUpdateQuery.keys(0), 1) = "$" Then 'update modifier

    Select Case UCase(oDicUpdateQuery.keys(0))
        Case UCase("$set")
            
            'get the dictionary with all fields to modify
            Dim oFieldsWithNewValue As Scripting.Dictionary
            Set oFieldsWithNewValue = oDicUpdateQuery.Item(oDicUpdateQuery.keys(0))
            
            Dim tmp_field_to_modify As Variant
            
            'passe en revue les documents
            For Each tmp_oid In oResultsSelect.documents.keys
                
                Dim oDocToModifyWithSetOperator As cls_NoSQL_Document
                Set oDocToModifyWithSetOperator = oResultsSelect.documents.Item(tmp_oid)
                
                'passe en revue les champs a modifier
                For Each tmp_field_to_modify In oFieldsWithNewValue.keys
                    
                    'deprecie pour assurer meilleure sync entre les 2 represenation dic/json des objets document
'                    If oDocToModifyWithSetOperator.representation_dictionary.Exists(tmp_field_to_modify) Then
'
'                        'object ou simple variable
'                        If VarType(oFieldsWithNewValue.Item(tmp_field_to_modify)) = vbObject Then
'                            Set oDocToModifyWithSetOperator.representation_dictionary.Item(tmp_field_to_modify) = oFieldsWithNewValue.Item(tmp_field_to_modify)
'                        Else
'                            oDocToModifyWithSetOperator.representation_dictionary.Item(tmp_field_to_modify) = oFieldsWithNewValue.Item(tmp_field_to_modify)
'                        End If
'
'
'                    Else
'
'                        oDocToModifyWithSetOperator.representation_dictionary.Add tmp_field_to_modify, oFieldsWithNewValue.Item(tmp_field_to_modify)
'
'                    End If
'
                    oDocToModifyWithSetOperator.edit_pair_key_value tmp_field_to_modify, oFieldsWithNewValue.Item(tmp_field_to_modify)
                Next
'
'                oDocToModifyWithSetOperator.sync_update_represenation_json_from_representation_dictionary
                
                'Debug.Print oDocToModifyWithSetOperator.representation_json
                
                update_deploy oJSON.toString(oDocToModifyWithSetOperator.representation_dictionary.Item("_id")), oDocToModifyWithSetOperator
                
                update.append_document oDocToModifyWithSetOperator
                
                If multi = False Then '1 update only
                    Exit For
                End If
                
            Next
            
        Case Else
            Debug.Print "ERROR cls_NoSQL_Collection::update unknow update modifier: " & oDicUpdateQuery.keys(0)
    End Select

Else 'full new doc
    
    Dim oDicFullNewDoc As Scripting.Dictionary
    Set oDicFullNewDoc = oUpdateQueryORFullNewDocument 'la query est en fait le nouveau document
    
    
    
    If oResultsSelect.count = 0 Then
        
        'insert 1
        insert oDicFullNewDoc
    Else
    
        For Each tmp_oid In oResultsSelect.documents.keys
            
            'append old ObjectId
            If oDicFullNewDoc.Exists("_id") Then
                oDicFullNewDoc.remove ("_id")
            Else
                
            End If
            
            
            Dim oNSDOld As cls_NoSQL_Document
                Set oNSDOld = oResultsSelect.documents.Item(tmp_oid)
            
            'check que "_id" est bien un sous-dico dans la representation_dictionary
            Dim check_dico_id As Scripting.Dictionary
            'Set check_dico_id = oNSDOld.representation_dictionary.Item("_id")
            oDicFullNewDoc.Add "_id", oNSDOld.representation_dictionary.Item("_id") 'dictionary object
            
            Dim oNSDFullNewDoc As cls_NoSQL_Document
                Set oNSDFullNewDoc = New cls_NoSQL_Document
            oNSDFullNewDoc.load_data oDicFullNewDoc
            oNSDFullNewDoc.link_with_collection Me
            
            'Debug.Print oJSON.toString(oNSDFullNewDoc.representation_dictionary.Item("_id"))
            
            update_deploy oJSON.toString(oNSDFullNewDoc.representation_dictionary.Item("_id")), oNSDFullNewDoc
            
            update.append_document oNSDFullNewDoc
            
            If multi = False Then '1 update only
                Exit For
            End If
        Next
    End If
End If



End Function


Private Function update_deploy(ByVal oid As String, oDoc As cls_NoSQL_Document)

update_deploy_in_collection oid, oDoc
update_deploy_in_sqlite_db oid, oDoc

End Function


Private Function update_deploy_in_collection(ByVal oid As String, oDoc As cls_NoSQL_Document)

If Me.documents.Exists(oid) = False Then
    Debug.Print "WARNING cls_NoSQL_Collection::update_deploy_in_collection _id not documents keys. Will create new one"
    Me.insert oDoc
Else
    
    's assure que disose d un _id
    If oDoc.representation_dictionary.Exists("_id") = False Then
        
        Debug.Print "WARNING cls_NoSQL_Collection::update_deploy_in_collection _id not present in document. Will append parameter oid"
        
        Dim oJSON As New jsonlib
        oDoc.representation_dictionary.Add "_id", oJSON.parse(oid) 'string doit revenir dic
        oDoc.representation_json = oJSON.toString(oDoc.representation_dictionary) 'sync avec autre representation
    End If
    
    Dim oDocOld As cls_NoSQL_Document
    'Set oDocOld = Me.documents.Item(oid) 'check
    
    Debug.Print "INFO cls_NoSQL_Collection::update_deploy_in_collection update document with _id: " & oid & " with new document: " & oDoc.representation_json
    
    Set Me.documents.Item(oid) = oDoc
    
End If

End Function


Private Function update_deploy_in_sqlite_db(ByVal oid As String, oDoc As cls_NoSQL_Document)

Dim sql_query As String
sql_query = "UPDATE " & Me.ref_database.get_sqlite_table_name_equivalent_from_collection_name(Me.name)
sql_query = sql_query & " SET " & Me.ref_database.sqlite_field_json_data & "=""" & Me.ref_database.nosql_encode_json_for_DB(oDoc.representation_json) & """"
sql_query = sql_query & " WHERE " & Me.ref_database.sqlite_field_oid & "=""" & Me.ref_database.nosql_encode_json_for_DB(oid) & """"

Debug.Print "INFO cls_NoSQL_Collection:update_deploy_in_sqlite_db with query: " & sql_query

sqlite3_query Me.ref_database.local_filepath, sql_query

End Function


Public Function remove(Optional oQueryDelete) As cls_NoSQL_QueryResult

Set remove = New cls_NoSQL_QueryResult

If IsMissing(oQueryDelete) Then
    Me.documents.RemoveAll
    remove_deploy_in_sqlite_db "", True
Else
    
    Dim oResultsSelect As cls_NoSQL_QueryResult
    Set oResultsSelect = Me.find(oQueryDelete) 'rapatrie les documents qui correspondent a la query de selection
    
    Dim tmp_oid As Variant
    For Each tmp_oid In oResultsSelect.documents.keys
        
        remove.append_document oResultsSelect.documents.Item(tmp_oid)
        
        Me.documents.remove tmp_oid
        remove_deploy_in_sqlite_db tmp_oid
        
    Next
    
End If

End Function


Private Function remove_deploy_in_sqlite_db(ByVal oid As String, Optional ByVal fullErase As Boolean = False)

Dim sql_query As String
sql_query = "DELETE FROM "
sql_query = sql_query & Me.ref_database.get_sqlite_table_name_equivalent_from_collection_name(Me.name)


If fullErase = False Then
    sql_query = sql_query & " WHERE " & Me.ref_database.sqlite_field_oid & "=""" & Me.ref_database.nosql_encode_json_for_DB(oid) & """"
End If

Debug.Print "INFO cls_NoSQL_Collection:remove_deploy_in_sqlite_db with query: " & sql_query

sqlite3_query Me.ref_database.local_filepath, sql_query

End Function


Public Function transform_query_into_object(oQuery) As Object

Dim oJSON As New jsonlib


If VarType(oQuery) = vbObject Then
    If TypeOf oQuery Is Scripting.Dictionary Then
        Set transform_query_into_object = oQuery
    ElseIf TypeOf oQuery Is Collection Then
        Set transform_query_into_object = oQuery
    Else
        Debug.Print "ERROR cls_NoSQL_Collection::transform_query_into_object no id about the kind of object for oQuery"
    End If
ElseIf VarType(oQuery) = vbString Then
    Set transform_query_into_object = oJSON.parse(CStr(oQuery))
Else
    Debug.Print "ERROR cls_NoSQL_Collection::transform_query_into_object no idea about oQuery's datatype"
End If

End Function


Public Function load_document(o, Optional ByVal oid As Variant) As cls_NoSQL_Document 's attend a ce qu il existe deja une cle

Set load_document = New cls_NoSQL_Document
load_document.link_with_collection Me
load_document.load_data o

If IsMissing(oid) = False Then
    Me.documents.Add oid, load_document
Else
    If load_document.representation_dictionary.Exists("_id") Then
        Me.documents.Add load_document.representation_dictionary.Item("_id"), load_document
    End If
End If

End Function


Private Sub Class_Terminate()

Set Me.documents = Nothing

End Sub
