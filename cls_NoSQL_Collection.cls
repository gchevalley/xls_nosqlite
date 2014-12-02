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

Set Me.documents = New Scripting.Dictionary 'pare a recevoir les objets cls_NoSQL_Document

Me.name = ""
Set Me.ref_database = New cls_NoSQL_Database

End Sub


' associe l objet collection avec la base de donnees qui l accueille
    ' @author : gchevalley
    ' @version : 0.1
    '
    ' @param oDB : objet cls_NoSQL_Database qui accueille physiquement les donnees de la collection
    '
    ' @return : void
Public Function link_with_database(oDB As cls_NoSQL_Database)

Debug.Print "INFO cls_NoSQL_Collection::link_with_database collection: " & Me.name & " with database: " & oDB.local_filepath

Set Me.ref_database = oDB

End Function



' insertion d un nouveau document dans la collection, s occupe d effectuer cette operation dans l objet Excel cls_NoSQL_Collection ainsi que dans la base de donnees physique
    ' @author : gchevalley
    ' @version : 0.1
    '
    ' @param o : le document qui peut etre un dictionnaire Excel ou de la notation JSON. Le premier niveau doit absolument etre un {objet} (pas un array, une collection ou un scalaire)
    ' [@param safe] by default False : apres l insertion, une seconde requete de selection pour s assurer que le document est bien present physiquement dans la base de donnees
    '
    ' @return : void
Public Function insert(o, Optional ByVal safe As Boolean = False) 'json string or excel's dictionary

Dim oJSON As New jsonlib

Dim tmp_document As cls_NoSQL_Document
    Set tmp_document = New cls_NoSQL_Document

tmp_document.load_data o
    

'append object id
If tmp_document.representation_dictionary.Exists("_id") Then
    Debug.Print "WARNING cls_NoSQL_Collection::insert document: " & tmp_document.representation_json & " already contains an _id, need to remove it"
    tmp_document.representation_dictionary.remove "_id"
End If

tmp_document.representation_dictionary.Add "_id", Me.ref_database.oOId.get_next_dic()


'regenere le doc pour assurer sync entre les 2 represenation
Dim final_document_to_store As cls_NoSQL_Document
    Set final_document_to_store = New cls_NoSQL_Document
final_document_to_store.load_data tmp_document.representation_dictionary

Set tmp_document = Nothing

'sync avec DB sqlite, la chaine json a besoin au prealable d un traitement pour pouvoir etre stocker dans la table sqlite
sqlite3_query Me.ref_database.local_filepath, "INSERT INTO " & Me.ref_database.get_sqlite_table_name_equivalent_from_collection_name(Me.name) & " (" & Me.ref_database.sqlite_field_oid & ", " & Me.ref_database.sqlite_field_json_data & ") VALUES (""" & Me.ref_database.nosql_encode_json_for_DB(oJSON.toString(final_document_to_store.representation_dictionary.Item("_id"))) & """,""" & Me.ref_database.nosql_encode_json_for_DB(final_document_to_store.representation_json) & """)"

If safe Then
    's assure que l entree est bien presente dans la DB
    Dim extract_sqlite As Variant
    extract_sqlite = sqlite3_query(Me.ref_database.local_filepath, "SELECT * FROM " & Me.ref_database.get_sqlite_table_name_equivalent_from_collection_name(Me.name) & " WHERE " & Me.ref_database.sqlite_field_oid & "=""" & Me.ref_database.nosql_encode_json_for_DB(oJSON.toString(final_document_to_store.representation_dictionary.Item("_id"))) & """")
    
    If UBound(extract_sqlite, 1) > 0 Then
    Else
        Debug.Print "ERROR cls_NoSQL_Collection::insert unable to insert document into database"
    End If
End If

'mise a jour de l object collection excel
Me.load_document final_document_to_store

End Function


' moteur de recherche de documents satisfaisants des conditions
    ' @author : gchevalley
    ' @version : 0.1
    '
    ' @param oQueryJSON : requete de selection qui peut etre un dictionnaire Excel ou de la notation JSON
    ' [@param limit_nbre_items] : borne superieure facultative du nombre d elements qui seront retournes
    ' [@param sort] : vecteur facultatif ordonne des champs sur lesquels le resultat doit etre trie, uniquement un seul niveau est supporte pour l instant, [field_1 : 1, field_2 : -1], 1 pour ASCENDING, -1 for DESCENDING
    ' [@param select_fields] : limite facultative les champs retournes {field_1: 1, field_2 : 1}, 1 signfique retourne le champ
    '
    ' @return : object cls_NoSQL_QueryResult
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

Dim tmp_constraint_field As Variant

Dim tmp_sub_constraint As Scripting.Dictionary
Dim tmp_key_sub_constraint As Variant 'devrait y en avoir qu une seule

If final_query_json Is Nothing Then
    Debug.Print "ERROR cls_NoSQL_Collection::find with oQueryJSON -> Quit"
    Exit Function
Else
    find.query = oJSON.toString(final_query_json)
    
    'check for sub queries
    If final_query_json.count = 1 Then
        
        For Each tmp_constraint_field In final_query_json.keys
        
            Select Case tmp_constraint_field
            
                Case "$or"
                
                    Dim col_or_sub_queries As collection
                    
                    If TypeOf final_query_json.Item(tmp_constraint_field) Is collection Then
                        
                        Set col_or_sub_queries = final_query_json.Item(tmp_constraint_field)
                        
                        Dim dic_or_sub_query As Object 'Scripting.Dictionary
                        
                        Dim oResultOr As cls_NoSQL_QueryResult
                        For Each dic_or_sub_query In col_or_sub_queries
                            
                            Set oResultOr = Me.find(dic_or_sub_query)
                            
                            Dim tmp_or_oid As Variant
                            For Each tmp_or_oid In oResultOr.documents.keys
                                If find.documents.Exists(tmp_or_oid) = False Then
                                    find.append_document oResultOr.documents.Item(tmp_or_oid)
                                Else
                                    Debug.Print "INFO cls_NoSQL_Collection::find _id: " & tmp_or_oid & " already in results"
                                End If
                            Next
                        Next
                        
                        Debug.Print "INFO cls_NoSQL_Collection::find query: " & oJSON.toString(final_query_json) & " returns: " & find.count & " result(s)"
                        Exit Function
                    Else
                        Debug.Print "ERROR cls_NoSQL_Collection::find $or need an array of queries objects : [{query1}, {query2}, ... , {queryn}]"
                    End If
                
                Case "$and"
                    
                    'transforme en query normal
                    Dim col_and_sub_queries As collection
                    
                    If TypeOf final_query_json.Item(tmp_constraint_field) Is collection Then
                        
                        Set col_and_sub_queries = final_query_json.Item(tmp_constraint_field)
                        
                        Dim dic_and_sub_query As Object 'Scripting.Dictionary
                        
                        Dim tmp_key_dic_and_sub_query As Variant
                        
                        Dim dic_normal_query As New Scripting.Dictionary
                        For Each dic_and_sub_query In col_and_sub_queries
                            
                            For Each tmp_key_dic_and_sub_query In dic_and_sub_query.keys
                                dic_normal_query.Add tmp_key_dic_and_sub_query, dic_and_sub_query.Item(tmp_key_dic_and_sub_query)
                            Next
                        Next
                        
                        Set find = Me.find(dic_normal_query)
                        
                        Debug.Print "INFO cls_NoSQL_Collection::find query: " & oJSON.toString(final_query_json) & " returns: " & find.count & " result(s)"
                        Exit Function
                    Else
                        Debug.Print "ERROR cls_NoSQL_Collection::find $or need an array of queries objects : [{query1}, {query2}, ... , {queryn}]"
                    End If
                
                Case "$nor"
                    
                    Dim dic_nor_or_query As New Scripting.Dictionary, dic_nor_and_query As New Scripting.Dictionary
                        dic_nor_or_query.Add "$or", final_query_json.Item(tmp_constraint_field)
                        dic_nor_and_query.Add "$and", final_query_json.Item(tmp_constraint_field)
                    
                    
                    Dim oQueryResultNorOr As cls_NoSQL_QueryResult, oQueryResultNorAnd As cls_NoSQL_QueryResult
                    Set oQueryResultNorOr = Me.find(dic_nor_or_query)
                        
                    Set oQueryResultNorAnd = Me.find(dic_nor_and_query)
                        
                    Dim tmp_nor_oid As Variant
                    For Each tmp_nor_oid In oQueryResultNorOr.documents.keys
                        If oQueryResultNorAnd.documents.Exists(tmp_nor_oid) = False Then
                            find.append_document oQueryResultNorOr.documents.Item(tmp_nor_oid)
                        End If
                    Next
                    
                    Debug.Print "INFO cls_NoSQL_Collection::find query: " & oJSON.toString(final_query_json) & " returns: " & find.count & " result(s)"
                    Exit Function
                    
            End Select

            Exit For
        Next
        
    End If
    
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
Dim final_sort_fields As collection 'attention l ordre compte -> ne doit pas etre un dico mais une collection
If IsMissing(sort) = False Then
    'set final_sort_fields
End If


'passe en revue les documents
Dim tmp_oid As Variant
Dim oDoc As cls_NoSQL_Document
Dim take_item As Boolean

Dim dic_constraints_find_candidate As New Scripting.Dictionary, dic_constraints_sub_conditions As Scripting.Dictionary

Dim tmp_scalar_value_in_array As Variant

Dim dic_check_in_constraints As Scripting.Dictionary
Dim tmp_in_scalar_value As Variant

For Each tmp_oid In Me.documents.keys ' parcours les documents
    
    Set oDoc = Me.documents.Item(tmp_oid)
    
    'passe en revue les contraintes
    Set dic_constraints_find_candidate = New Scripting.Dictionary
    
    For Each tmp_constraint_field In final_query_json.keys 'parcours les champs contraints
        
        'type de contrainte scalaire, complexe -> egalement checker operator, optimiser avec regexp
        Dim is_query_operator As Boolean
        is_query_operator = False
        If (VarType(final_query_json.Item(tmp_constraint_field)) = vbObject And TypeOf final_query_json.Item(tmp_constraint_field) Is Scripting.Dictionary) Then
            
            'transforme en dico si necessaire
            If (VarType(final_query_json.Item(tmp_constraint_field)) = vbObject And TypeOf final_query_json.Item(tmp_constraint_field) Is Scripting.Dictionary) Then
                Set tmp_sub_constraint = final_query_json.Item(tmp_constraint_field)
            Else
                Set tmp_sub_constraint = oJSON.parse(CStr(final_query_json.Item(tmp_constraint_field)))
            End If
            
            is_query_operator = True
            
        Else
            
        End If
        
        
        Dim col_objects_to_check As collection 'un obet candidat contient juste tout le schema de cles necessaire et se trouve au dernier niveau il ne reste qu a checker la contrainte
        Set col_objects_to_check = oDoc.get_collection_sub_objects(tmp_constraint_field) 'collection d objet dictionaire ou collection d une collection (array)
        
        Dim dic_CheckLevelConstraint As Scripting.Dictionary ', dic_current_level As Scripting.Dictionary, dic_sub_level As Scripting.Dictionary
        
        
        'Set dic_constraints_sub_conditions = New Scripting.Dictionary
        
        If col_objects_to_check.count > 0 Then
            
            Dim tmp_constraint_field_last_level As String
            tmp_constraint_field_last_level = find_get_last_level_on_query_field(tmp_constraint_field)
            
            Dim tmp_sub_dic_or_col As Object
            For Each tmp_sub_dic_or_col In col_objects_to_check 'parcours les objets du document
                
                Set dic_constraints_sub_conditions = New Scripting.Dictionary
                
                If TypeOf tmp_sub_dic_or_col Is Scripting.Dictionary Then
                        
                    Set dic_CheckLevelConstraint = tmp_sub_dic_or_col
                    
                    If dic_CheckLevelConstraint.Exists(tmp_constraint_field_last_level) Then
                        
                        'type de contrainte scalaire, complexe -> egalement checker operator, optimiser avec regexp
                        If is_query_operator Then
                            
                            'repere l operator
                            For Each tmp_key_sub_constraint In tmp_sub_constraint.keys 'un meme champs peut avoir plusieurs conditions
                                
                                dic_constraints_sub_conditions.Add tmp_key_sub_constraint, False ' pour chaque condition d un meme champ
                                
                                If dic_constraints_sub_conditions.Item(tmp_key_sub_constraint) = False Then
                                
                                    Select Case tmp_key_sub_constraint
                                    
                                        Case "$gt"
                                            If IsNumeric(tmp_sub_constraint.Item(tmp_key_sub_constraint)) Then
                                                If IsNumeric(dic_CheckLevelConstraint.Item(tmp_constraint_field_last_level)) Then
                                                    If dic_CheckLevelConstraint.Item(tmp_constraint_field_last_level) > tmp_sub_constraint.Item(tmp_key_sub_constraint) Then
                                                        dic_constraints_sub_conditions.Item(tmp_key_sub_constraint) = True
                                                    End If
                                                End If
                                            Else
                                                If dic_CheckLevelConstraint.Item(tmp_constraint_field_last_level) > tmp_sub_constraint.Item(tmp_key_sub_constraint) Then
                                                    dic_constraints_sub_conditions.Item(tmp_key_sub_constraint) = True
                                                End If
                                            End If
                                        
                                        
                                        Case "$gte"
                                            If IsNumeric(tmp_sub_constraint.Item(tmp_key_sub_constraint)) Then
                                                If IsNumeric(dic_CheckLevelConstraint.Item(tmp_constraint_field_last_level)) Then
                                                    If dic_CheckLevelConstraint.Item(tmp_constraint_field_last_level) >= tmp_sub_constraint.Item(tmp_key_sub_constraint) Then
                                                        dic_constraints_sub_conditions.Item(tmp_key_sub_constraint) = True
                                                    End If
                                                End If
                                            Else
                                                If dic_CheckLevelConstraint.Item(tmp_constraint_field_last_level) >= tmp_sub_constraint.Item(tmp_key_sub_constraint) Then
                                                    dic_constraints_sub_conditions.Item(tmp_key_sub_constraint) = True
                                                End If
                                            End If
                                        
                                        
                                        Case "$lt"
                                            If IsNumeric(tmp_sub_constraint.Item(tmp_key_sub_constraint)) Then
                                                If IsNumeric(dic_CheckLevelConstraint.Item(tmp_constraint_field_last_level)) Then
                                                    If dic_CheckLevelConstraint.Item(tmp_constraint_field_last_level) < tmp_sub_constraint.Item(tmp_key_sub_constraint) Then
                                                        dic_constraints_sub_conditions.Item(tmp_key_sub_constraint) = True
                                                    End If
                                                End If
                                            Else
                                                If dic_CheckLevelConstraint.Item(tmp_constraint_field_last_level) < tmp_sub_constraint.Item(tmp_key_sub_constraint) Then
                                                    dic_constraints_sub_conditions.Item(tmp_key_sub_constraint) = True
                                                End If
                                            End If
                                        
                                        
                                        Case "$lte"
                                            If IsNumeric(tmp_sub_constraint.Item(tmp_key_sub_constraint)) Then
                                                If IsNumeric(dic_CheckLevelConstraint.Item(tmp_constraint_field_last_level)) Then
                                                    If dic_CheckLevelConstraint.Item(tmp_constraint_field_last_level) <= tmp_sub_constraint.Item(tmp_key_sub_constraint) Then
                                                        dic_constraints_sub_conditions.Item(tmp_key_sub_constraint) = True
                                                    End If
                                                End If
                                            Else
                                                If dic_CheckLevelConstraint.Item(tmp_constraint_field_last_level) <= tmp_sub_constraint.Item(tmp_key_sub_constraint) Then
                                                    dic_constraints_sub_conditions.Item(tmp_key_sub_constraint) = True
                                                End If
                                            End If
                                        
                                        
                                        Case "$ne"
                                            If IsNumeric(tmp_sub_constraint.Item(tmp_key_sub_constraint)) Then
                                                If IsNumeric(dic_CheckLevelConstraint.Item(tmp_constraint_field_last_level)) Then
                                                    If dic_CheckLevelConstraint.Item(tmp_constraint_field_last_level) <> tmp_sub_constraint.Item(tmp_key_sub_constraint) Then
                                                        dic_constraints_sub_conditions.Item(tmp_key_sub_constraint) = True
                                                    End If
                                                End If
                                            Else
                                                If dic_CheckLevelConstraint.Item(tmp_constraint_field_last_level) <> tmp_sub_constraint.Item(tmp_key_sub_constraint) Then
                                                    dic_constraints_sub_conditions.Item(tmp_key_sub_constraint) = True
                                                End If
                                            End If
                                        
                                        
                                        Case "$exists"
                                            dic_constraints_sub_conditions.Item(tmp_key_sub_constraint) = True
                                        
                                        
                                        Case "$in"
                                            
                                            Set dic_check_in_constraints = New Scripting.Dictionary
                                    
                                            For Each tmp_in_scalar_value In tmp_sub_constraint.Item(tmp_key_sub_constraint)
                                                
                                                dic_check_in_constraints.Add tmp_in_scalar_value, False
                                                
                                                If dic_CheckLevelConstraint.Item(tmp_constraint_field_last_level) = tmp_in_scalar_value Then
                                                    dic_check_in_constraints.Item(tmp_in_scalar_value) = True
                                                    Exit For
                                                End If
                                                
                                            Next
                                            
                                            dic_constraints_sub_conditions.Item(tmp_key_sub_constraint) = False
                                            For Each tmp_in_scalar_value In dic_check_in_constraints.keys
                                                If dic_check_in_constraints.Item(tmp_in_scalar_value) = True Then
                                                    dic_constraints_sub_conditions.Item(tmp_key_sub_constraint) = True
                                                End If
                                            Next
                                        
                                        
                                        Case "$nin"
                                            
                                            Set dic_check_in_constraints = New Scripting.Dictionary
                                    
                                            For Each tmp_in_scalar_value In tmp_sub_constraint.Item(tmp_key_sub_constraint)
                                                
                                                dic_check_in_constraints.Add tmp_in_scalar_value, False
                                                
                                                If dic_CheckLevelConstraint.Item(tmp_constraint_field_last_level) = tmp_in_scalar_value Then
                                                    dic_check_in_constraints.Item(tmp_in_scalar_value) = True
                                                    Exit For
                                                End If
                                                
                                            Next
                                            
                                            dic_constraints_sub_conditions.Item(tmp_key_sub_constraint) = True
                                            For Each tmp_in_scalar_value In dic_check_in_constraints.keys
                                                If dic_check_in_constraints.Item(tmp_in_scalar_value) = True Then
                                                    dic_constraints_sub_conditions.Item(tmp_key_sub_constraint) = False
                                                End If
                                            Next
                                            
                                        
                                        Case "$regex"
                                            
                                            oReg.Pattern = tmp_sub_constraint.Item(tmp_key_sub_constraint)
                                            
                                            Set matches = oReg.Execute(dic_CheckLevelConstraint.Item(tmp_constraint_field_last_level))
                                            
                                            If matches.count > 0 Then
                                                dic_constraints_sub_conditions.Item(tmp_key_sub_constraint) = True
                                            End If
                                            
                                        
                                        Case Else
                                            Debug.Print "WARNING cls_NoSQL_Collection::find unkown operator " & tmp_key_sub_constraint
                                    End Select
                                Else
                                    'la sous-condition du champs est deja respectée
                                End If
                                
                            Next 'boucle sur chaque sous - contrainte de chaque champs, car un meme champs peut avoir plusieurs conditions
                        
                        Else 'n est pas un query operator -> simple egality check
                            
                            ' simple check equality sur cle d une seul dictionnaire
                            dic_constraints_sub_conditions.Add tmp_constraint_field, False ' pour chaque condition d un meme champ
                            If dic_CheckLevelConstraint.Item(tmp_constraint_field_last_level) = final_query_json.Item(tmp_constraint_field) Then
                                dic_constraints_sub_conditions.Item(tmp_constraint_field) = True
                            End If
                            
                        End If
                        
                        
                    End If
                
                ElseIf TypeOf tmp_sub_dic_or_col Is collection Then
                    
                    
                    If is_query_operator Then
                        
                        'repere l operator
                        For Each tmp_key_sub_constraint In tmp_sub_constraint.keys
                            
                            dic_constraints_sub_conditions.Add tmp_key_sub_constraint, False
                            
                            Select Case tmp_key_sub_constraint
                                
                                Case "$in" 'at least one...
                                    
                                    Set dic_check_in_constraints = New Scripting.Dictionary
                                    
                                    For Each tmp_in_scalar_value In tmp_sub_constraint.Item(tmp_key_sub_constraint)
                                        
                                        dic_check_in_constraints.Add tmp_in_scalar_value, False
                                        
                                        For Each tmp_scalar_value_in_array In tmp_sub_dic_or_col
                                            If tmp_scalar_value_in_array = tmp_in_scalar_value Then
                                                dic_check_in_constraints.Item(tmp_in_scalar_value) = True
                                                Exit For
                                            End If
                                        Next
                                        
                                    Next
                                    
                                    For Each tmp_in_scalar_value In dic_check_in_constraints.keys
                                        If dic_check_in_constraints.Item(tmp_in_scalar_value) = True Then
                                            dic_constraints_sub_conditions.Item(tmp_key_sub_constraint) = True
                                        End If
                                    Next
                                
                                Case "$nin"
                                    
                                    Set dic_check_in_constraints = New Scripting.Dictionary
                                    
                                    For Each tmp_in_scalar_value In tmp_sub_constraint.Item(tmp_key_sub_constraint)
                                        
                                        dic_check_in_constraints.Add tmp_in_scalar_value, False
                                        
                                        For Each tmp_scalar_value_in_array In tmp_sub_dic_or_col
                                            If tmp_scalar_value_in_array = tmp_in_scalar_value Then
                                                dic_check_in_constraints.Item(tmp_in_scalar_value) = True
                                                Exit For
                                            End If
                                        Next
                                        
                                    Next
                                    
                                    dic_constraints_sub_conditions.Item(tmp_key_sub_constraint) = True
                                    For Each tmp_in_scalar_value In dic_check_in_constraints.keys
                                        If dic_check_in_constraints.Item(tmp_in_scalar_value) = True Then
                                            dic_constraints_sub_conditions.Item(tmp_key_sub_constraint) = False
                                        End If
                                    Next
                                
                                Case "$all" 'need all
                                    
                                    Set dic_check_in_constraints = New Scripting.Dictionary
                                    
                                    For Each tmp_in_scalar_value In tmp_sub_constraint.Item(tmp_key_sub_constraint)
                                        
                                        dic_check_in_constraints.Add tmp_in_scalar_value, False
                                        
                                        For Each tmp_scalar_value_in_array In tmp_sub_dic_or_col
                                            If tmp_scalar_value_in_array = tmp_in_scalar_value Then
                                                dic_check_in_constraints.Item(tmp_in_scalar_value) = True
                                                Exit For
                                            End If
                                        Next
                                        
                                    Next
                                    
                                    dic_constraints_sub_conditions.Item(tmp_key_sub_constraint) = True
                                    For Each tmp_in_scalar_value In dic_check_in_constraints.keys
                                        If dic_check_in_constraints.Item(tmp_in_scalar_value) = False Then
                                            dic_constraints_sub_conditions.Item(tmp_key_sub_constraint) = False
                                        End If
                                    Next
                                    
                                
                                Case "$size"
                                        
                                    dic_constraints_sub_conditions.Item(tmp_key_sub_constraint) = False
                                    
                                    If IsNumeric(tmp_sub_constraint.Item(tmp_key_sub_constraint)) Then
                                        If tmp_sub_dic_or_col.count = tmp_sub_constraint.Item(tmp_key_sub_constraint) Then
                                            dic_constraints_sub_conditions.Item(tmp_key_sub_constraint) = True
                                        End If
                                    End If

                                    
                                
                                Case Else
                                    Debug.Print "WARNING cls_NoSQL_Collection::find unkown operator " & tmp_key_sub_constraint
                            End Select
                        
                        Next
                        
                    Else
                        
                        Debug.Print "ERROR cls_NoSQL_Collection::find query on array field without a query operator ($all, $in, $nin, $size)"
                        
                    End If
                    
                
                End If
                
                
                If dic_constraints_find_candidate.Exists(tmp_constraint_field) = False Then
                    dic_constraints_find_candidate.Add tmp_constraint_field, dic_constraints_sub_conditions 'pour un champs, ajoute l etat de toutes les sous-conditions
                Else
                    
                    'le sous-objet que l on vient de controler respecte - il toutes les sous contraintes ?
                    Dim find_false_sub_object_constraints As Boolean
                        find_false_sub_object_constraints = False
                    For Each tmp_key_sub_constraint In dic_constraints_sub_conditions
                        If dic_constraints_sub_conditions.Item(tmp_key_sub_constraint) = False Then
                            find_false_sub_object_constraints = True
                        End If
                    Next
                    
                    If find_false_sub_object_constraints = False Then
                        'replace
                        Set dic_constraints_find_candidate.Item(tmp_constraint_field) = dic_constraints_sub_conditions
                    End If
                    
                End If
                
                
            Next 'loop candidate dans (sous-) objet ?
            
        Else ' aucun objet dans la collection
            
            'dic_constraints_sub_conditions n a pas ete initialise
            Set dic_constraints_sub_conditions = New Scripting.Dictionary
            dic_constraints_sub_conditions.Add tmp_key_sub_constraint, False
        
        End If 'y a t il au moins un sous-objet ?
        
        
        'dic_constraints_find_candidate.Add tmp_constraint_field, dic_constraints_sub_conditions 'pour un champs, ajoute l etat de toutes les sous-conditions
        
    Next 'loop field constraint
    
    
    'toutes les contraintes sont elles respectee ?
    Dim tmp_constraints As Variant
    take_item = True
    If dic_constraints_find_candidate.count > 0 Then
        For Each tmp_constraint_field In dic_constraints_find_candidate.keys
            
            'toute les sous conditions sont elles bien a true ?
            Set dic_constraints_sub_conditions = dic_constraints_find_candidate.Item(tmp_constraint_field)
            
            If dic_constraints_sub_conditions.count = 0 Or dic_constraints_sub_conditions Is Nothing Then
                take_item = False
                Exit For
            Else
                For Each tmp_key_sub_constraint In dic_constraints_sub_conditions
                    If take_item = True And dic_constraints_sub_conditions.Item(tmp_key_sub_constraint) = False Then
                        take_item = False
                        Exit For
                    End If
                Next
            End If
        Next
    Else
        take_item = False
    End If
    
    
check_state_take_item:
    If take_item Then
        
        If IsMissing(select_fields) = True Then
add_full_doc_to_result:
            
            find.append_document oDoc
        Else
            
            If checked_select_fields = True Then
            
                Dim tmp_dic_with_filtered_fields As Scripting.Dictionary
                    Set tmp_dic_with_filtered_fields = New Scripting.Dictionary
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


' pour les requetes sur les sous-objets d un document, permet de retourner la cle du tout dernier niveau
    ' @author : gchevalley
    ' @version : 0.1
    '
    ' @param query_field : le champ de la requete complet : top_level_key.second_level_key.[...].[last_level_key]
    '
    ' @return : String du nom de la cle du dernier niveau (last_level_key)
Private Function find_get_last_level_on_query_field(ByVal query_field As String) As String

If InStr(query_field, ".") <> 0 Then
    find_get_last_level_on_query_field = StrReverse(Left(StrReverse(query_field), InStr(StrReverse(query_field), ".") - 1))
Else
    find_get_last_level_on_query_field = query_field
End If

End Function


' requete de mise a jour de documents
    ' @author : gchevalley
    ' @version : 0.1
    '
    ' @param oSelectQuery : requete de selection des documents a modifier a l aide d un dictionnaire Excel ou de la notation JSON
    ' @param oUpdateQueryORFullNewDocument : nouveau document qui doit ecraser le precedent ou une requete d edition de certaines parties d un document
    ' [@param multi] by default False : edite uniquement le premier document trouve ou la totalite des documents qui correspondent a la requete de recherche
    '
    ' @return : objet cls_NoSQL_QueryResult contenant les cls_NoSQL_Document ayant recu une modification
Public Function update(oSelectQuery, oUpdateQueryORFullNewDocument, Optional ByVal multi As Boolean = False, Optional ByVal safe As Boolean = False) As cls_NoSQL_QueryResult

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

If oDicUpdateQuery Is Nothing Then
    Debug.Print "ERROR cls_NoSQL_Collection::update objectoUpdateQueryORFullNewDocument not valid ! -> Exit."
    Exit Function
End If


'distinct full new doc (erase old replace with new) or update modifier (just part(s) modifier)
Dim tmp_oid As Variant
If oDicUpdateQuery.count = 1 And Left(oDicUpdateQuery.keys(0), 1) = "$" Then 'update modifier

    Select Case UCase(oDicUpdateQuery.keys(0))
        
        Dim oFieldsWithNewValue As Scripting.Dictionary
        Dim oDocToModifyWithSetOperator As cls_NoSQL_Document
        
        Dim oKeyToDelete As Scripting.Dictionary
        Dim oDocToModifyWithUnsetOperator As cls_NoSQL_Document
        
        'on fields
        Case UCase("$set") 'edit content for keys
            
            'get the dictionary with all fields to modify
            Set oFieldsWithNewValue = oDicUpdateQuery.Item(oDicUpdateQuery.keys(0))
            
            Dim tmp_field_to_modify As Variant
            
            'passe en revue les documents
            For Each tmp_oid In oResultsSelect.documents.keys
                
                Set oDocToModifyWithSetOperator = oResultsSelect.documents.Item(tmp_oid)
                
                'passe en revue les champs a modifier
                For Each tmp_field_to_modify In oFieldsWithNewValue.keys
                    oDocToModifyWithSetOperator.edit_pair_key_value tmp_field_to_modify, oFieldsWithNewValue.Item(tmp_field_to_modify)
                Next
                
                'Debug.Print oDocToModifyWithSetOperator.representation_json
                
                update_deploy oJSON.toString(oDocToModifyWithSetOperator.representation_dictionary.Item("_id")), oDocToModifyWithSetOperator, safe
                
                update.append_document oDocToModifyWithSetOperator
                
                If multi = False Then '1 update only
                    Exit For
                End If
                
            Next
        
        
        Case UCase("$unset") 'detruit une cle ainsi que son contenu
            
            Set oKeyToDelete = oDicUpdateQuery.Item(oDicUpdateQuery.keys(0))
            
            Dim tmp_key_to_delete As Variant
            
            'passe en revue les documents
            For Each tmp_oid In oResultsSelect.documents.keys
            
                Set oDocToModifyWithUnsetOperator = oResultsSelect.documents.Item(tmp_oid)
                
                'passe en revue les keys a detruire
                For Each tmp_key_to_delete In oKeyToDelete.keys
                    oDocToModifyWithUnsetOperator.remove_key tmp_key_to_delete
                Next
                
                
                update_deploy oJSON.toString(oDocToModifyWithUnsetOperator.representation_dictionary.Item("_id")), oDocToModifyWithUnsetOperator, safe
                
                update.append_document oDocToModifyWithUnsetOperator
                
                If multi = False Then '1 update only
                    Exit For
                End If
                
                
            Next
        
        Case UCase("$rename")
            
            Dim oKeysToRename As Scripting.Dictionary
            Set oKeysToRename = oDicUpdateQuery.Item(oDicUpdateQuery.keys(0))
            
            Dim tmp_key_to_rename As Variant
            
            'passe en revue les documents
            Dim oDocToModifyWithRenameOperator As cls_NoSQL_Document
            For Each tmp_oid In oResultsSelect.documents.keys
                Set oDocToModifyWithRenameOperator = oResultsSelect.documents.Item(tmp_oid)
                
                'passe en revue les keys a detruire
                For Each tmp_key_to_rename In oKeysToRename.keys
                    oDocToModifyWithRenameOperator.rename_key tmp_key_to_rename, oKeysToRename.Item(tmp_key_to_rename)
                Next
                
                
                update_deploy oJSON.toString(oDocToModifyWithRenameOperator.representation_dictionary.Item("_id")), oDocToModifyWithRenameOperator, safe
                
                update.append_document oDocToModifyWithRenameOperator
                
                If multi = False Then '1 update only
                    Exit For
                End If
                
            Next
            
        
        Case UCase("$max")
            
            Dim oKeysToMax As Scripting.Dictionary
            Set oKeysToMax = oDicUpdateQuery.Item(oDicUpdateQuery.keys(0))
            
            Dim tmp_key_to_max As Variant
            
            'passe en revue les documents
            Dim oDocToModifyWithMaxOperator As cls_NoSQL_Document
            For Each tmp_oid In oResultsSelect.documents.keys
                Set oDocToModifyWithMaxOperator = oResultsSelect.documents.Item(tmp_oid)
                
                'passe en revue les keys a maximuser
                For Each tmp_key_to_max In oKeysToMax.keys
                    
                    If IsNumeric(oKeysToMax.Item(tmp_key_to_max)) Then
                        If oDocToModifyWithMaxOperator.min_max_value_for_one_key(">", tmp_key_to_max, oKeysToMax.Item(tmp_key_to_max)) Then
                            update_deploy oJSON.toString(oDocToModifyWithMaxOperator.representation_dictionary.Item("_id")), oDocToModifyWithMaxOperator, safe
                            update.append_document oDocToModifyWithMaxOperator
                        End If
                        
                    Else
                        Debug.Print "cls_NoSQL_Collection::update $max new value: " & oKeysToMax.Item(tmp_key_to_max) & " is not numerical"
                    End If
                Next
                
                
                If multi = False Then '1 update only
                    Exit For
                End If
                
            Next
        
        
        Case UCase("$min")
            
            Dim oKeysToMin As Scripting.Dictionary
            Set oKeysToMin = oDicUpdateQuery.Item(oDicUpdateQuery.keys(0))
            
            Dim tmp_key_to_min As Variant
            
            'passe en revue les documents
            Dim oDocToModifyWithMinOperator As cls_NoSQL_Document
            For Each tmp_oid In oResultsSelect.documents.keys
                Set oDocToModifyWithMinOperator = oResultsSelect.documents.Item(tmp_oid)
                
                'passe en revue les keys a maximuser
                For Each tmp_key_to_min In oKeysToMin.keys
                    
                    If IsNumeric(oKeysToMin.Item(tmp_key_to_min)) Then
                        If oDocToModifyWithMinOperator.min_max_value_for_one_key("<", tmp_key_to_min, oKeysToMin.Item(tmp_key_to_min)) Then
                            update_deploy oJSON.toString(oDocToModifyWithMinOperator.representation_dictionary.Item("_id")), oDocToModifyWithMinOperator, safe
                            update.append_document oDocToModifyWithMinOperator
                        End If
                        
                    Else
                        Debug.Print "cls_NoSQL_Collection::update $min new value: " & oKeysToMin.Item(tmp_key_to_min) & " is not numerical"
                    End If
                Next
                
                
                If multi = False Then '1 update only
                    Exit For
                End If
                
            Next
            
            
        Case UCase("$inc")
            
            Dim oKeysToInc As Scripting.Dictionary
            Set oKeysToInc = oDicUpdateQuery.Item(oDicUpdateQuery.keys(0))
            
            Dim tmp_key_to_inc As Variant
            
            'passe en revue les documents
            Dim oDocToModifyWithIncOperator As cls_NoSQL_Document
            For Each tmp_oid In oResultsSelect.documents.keys
                Set oDocToModifyWithIncOperator = oResultsSelect.documents.Item(tmp_oid)
                
                'passe en revue les keys a maximuser
                For Each tmp_key_to_inc In oKeysToInc.keys
                    
                    If IsNumeric(oKeysToInc.Item(tmp_key_to_inc)) Then
                        If oDocToModifyWithIncOperator.dephasing_value_for_one_key(tmp_key_to_inc, oKeysToInc.Item(tmp_key_to_inc)) Then
                            update_deploy oJSON.toString(oDocToModifyWithIncOperator.representation_dictionary.Item("_id")), oDocToModifyWithIncOperator, safe
                            update.append_document oDocToModifyWithIncOperator
                        End If
                        
                    Else
                        Debug.Print "cls_NoSQL_Collection::update $inc new value: " & oKeysToInc.Item(tmp_key_to_inc) & " is not numerical"
                    End If
                Next
                
                
                If multi = False Then '1 update only
                    Exit For
                End If
                
            Next
            
        
        Case UCase("$push")
            
            's assure que le contenu en face de la cle est bien un array/collection
            Dim oKeysToPush As Scripting.Dictionary
            Set oKeysToPush = oDicUpdateQuery.Item(oDicUpdateQuery.keys(0))
            
            Dim tmp_key_to_push As Variant
            
            'passe en revue les documents
            Dim oDocToModifyWithPushOperator As cls_NoSQL_Document
            For Each tmp_oid In oResultsSelect.documents.keys
                Set oDocToModifyWithPushOperator = oResultsSelect.documents.Item(tmp_oid)
                
                'passe en revue les keys a maximuser
                For Each tmp_key_to_push In oKeysToPush.keys
                    If oDocToModifyWithPushOperator.append_value_to_an_array(tmp_key_to_push, oKeysToPush.Item(tmp_key_to_push)) Then
                        update_deploy oJSON.toString(oDocToModifyWithPushOperator.representation_dictionary.Item("_id")), oDocToModifyWithPushOperator, safe
                        update.append_document oDocToModifyWithPushOperator
                    End If
                Next
                
                
                If multi = False Then '1 update only
                    Exit For
                End If
                
            Next
            
        
        Case UCase("$pop")
            
            's assure que le contenu en face de la cle est bien un array/collection
            Dim oKeysToPop As Scripting.Dictionary
            Set oKeysToPop = oDicUpdateQuery.Item(oDicUpdateQuery.keys(0))
            
            Dim tmp_key_to_pop As Variant
            
            'passe en revue les documents
            Dim oDocToModifyWithPopOperator As cls_NoSQL_Document
            For Each tmp_oid In oResultsSelect.documents.keys
                Set oDocToModifyWithPopOperator = oResultsSelect.documents.Item(tmp_oid)
                
                'passe en revue les keys a maximuser
                For Each tmp_key_to_pop In oKeysToPop.keys
                    If oDocToModifyWithPopOperator.pop_element_from_an_array(tmp_key_to_pop, oKeysToPop.Item(tmp_key_to_pop)) Then
                        update_deploy oJSON.toString(oDocToModifyWithPopOperator.representation_dictionary.Item("_id")), oDocToModifyWithPopOperator, safe
                        update.append_document oDocToModifyWithPopOperator
                    End If
                Next
                
                
                If multi = False Then '1 update only
                    Exit For
                End If
                
            Next
            
        
        Case UCase("$pull")
            
            's assure que le contenu en face de la cle est bien un array/collection
            Dim oKeysToPull As Scripting.Dictionary
            Set oKeysToPull = oDicUpdateQuery.Item(oDicUpdateQuery.keys(0))
            
            Dim tmp_key_to_pull As Variant
            
            'passe en revue les documents
            Dim oDocToModifyWithPullOperator As cls_NoSQL_Document
            For Each tmp_oid In oResultsSelect.documents.keys
                Set oDocToModifyWithPullOperator = oResultsSelect.documents.Item(tmp_oid)
                
                'passe en revue les keys a maximuser
                For Each tmp_key_to_pull In oKeysToPull.keys
                    If oDocToModifyWithPullOperator.remove_element_from_an_array_based_on_content(tmp_key_to_pull, oKeysToPull.Item(tmp_key_to_pull)) Then
                        update_deploy oJSON.toString(oDocToModifyWithPullOperator.representation_dictionary.Item("_id")), oDocToModifyWithPullOperator, safe
                        update.append_document oDocToModifyWithPullOperator
                    End If
                Next
                
                
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


' lorsqu un document d une collection a ete modifie, il est necessaire d appliquer/repliquer les changements sur l objet cls_NoSQL_Collection ainsi que dans la base de donnees physique afin que ces differents contenus restent bien synchronises
    ' @author : gchevalley
    ' @version : 0.1
    '
    ' @param oid : l identifiant unique du document sous sa forme texte (JSON) qui correspond a la cle unique du dictionnaire qui stocke les cls_NoSQL_Document dans les objets cls_NoSQL_Collection ainsi que le champs id de la base de donnees physique
    ' @param oDoc : le document dans sa nouvelle forme complete
    ' [@param safe] by default False : apres l insertion, une seconde requete de selection pour s assurer que le document est bien present physiquement dans la base de donnees
    '
    ' @return : void
Private Function update_deploy(ByVal oid As String, oDoc As cls_NoSQL_Document, Optional ByVal safe As Boolean = False)

update_deploy_in_collection oid, oDoc
update_deploy_in_sqlite_db oid, oDoc, safe

End Function


' Applique la mise a jour d un document dans l objet cls_NoSQL_Collection
' si aucun document avec l identifiant n est trouve, une nouvelle insertion a lieu
    ' @author : gchevalely
    ' @version : 0.1
    '
    ' @param oid : l identifiant unique du document sous sa forme texte (JSON) qui correspond a la cle unique du dictionnaire qui stocke les cls_NoSQL_Document dans les objets cls_NoSQL_Collection ainsi que le champs id de la base de donnees physique
    ' @param oDoc : le document modifie dans sa nouvelle forme complete
    '
    ' @return : void
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


' Applique les modifications d un document dans la base de donnees physique
    ' @author : gchevalley
    ' @version : 0.1
    '
    ' @param oid : l identifiant unique du document dans la base de donnees physique
    ' @param oDoc : le document modifie dans sa nouvelle forme complete
    ' [@param safe] by default False : apres l insertion, une seconde requete de selection pour s assurer que le document est bien present physiquement dans la base de donnees
    '
    ' @return : void
Private Function update_deploy_in_sqlite_db(ByVal oid As String, oDoc As cls_NoSQL_Document, Optional ByVal safe As Boolean = False)

Dim sql_query As String
sql_query = "UPDATE " & Me.ref_database.get_sqlite_table_name_equivalent_from_collection_name(Me.name)
sql_query = sql_query & " SET " & Me.ref_database.sqlite_field_json_data & "=""" & Me.ref_database.nosql_encode_json_for_DB(oDoc.representation_json) & """"
sql_query = sql_query & " WHERE " & Me.ref_database.sqlite_field_oid & "=""" & Me.ref_database.nosql_encode_json_for_DB(oid) & """"

Debug.Print "INFO cls_NoSQL_Collection:update_deploy_in_sqlite_db with query: " & sql_query

If safe = True Then
    Dim extract_old_content As Variant
    extract_old_content = sqlite3_query(Me.ref_database.local_filepath, "SELECT " & Me.ref_database.sqlite_field_json_data & " FROM " & Me.ref_database.get_sqlite_table_name_equivalent_from_collection_name(Me.name) & " WHERE " & Me.ref_database.sqlite_field_oid & "=""" & Me.ref_database.nosql_encode_json_for_DB(oid) & """")
End If

sqlite3_query Me.ref_database.local_filepath, sql_query


If safe = True Then
    If UBound(extract_old_content, 1) > 0 Then ' an entry was already existing
        
        Dim extract_new_content As Variant
        extract_new_content = sqlite3_query(Me.ref_database.local_filepath, "SELECT " & Me.ref_database.sqlite_field_json_data & " FROM " & Me.ref_database.get_sqlite_table_name_equivalent_from_collection_name(Me.name) & " WHERE " & Me.ref_database.sqlite_field_oid & "=""" & Me.ref_database.nosql_encode_json_for_DB(oid) & """")
        
        
        If extract_old_content(1)(0) <> extract_new_content(1)(0) Then
        Else
            Debug.Print "ERROR / WARNING cls_NoSQL_Collection::update_deploy_in_sqlite_db Same content before and after the update in database. Unable to update ?"
        End If
        
    End If
End If


End Function


' suppression de documents d une collection sur la base d une requete de selection
' une fonction separee se charge de repliquer le changement dans la base de donnees physique pour assurer la synchronisation des contenus
    ' @author : gchevalley
    ' @version : 0.1
    '
    ' @param oQueryDelete : requete de selection des documents a supprimer. Formats admis : Dictionnaire Excel ou notation JSON
    '
    ' @return : objet cls_NoSQL_QueryResult contenant les cls_NoSQL_Document supprimes
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


' replication d une suppresion d un document dans la base de donnees physique
    ' @author : gchevalley
    ' @version : 0.1
    '
    ' @param oid : l identifiant unique du document dans la base de donnees physique, cette requete peut etre vide si le second parametre de cette fonction est a True
    ' [@param fullErase], by default False, suppresion de la totalite des enregistrements de la table correspondant a la collection
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
    ElseIf TypeOf oQuery Is collection Then
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


' [ {$match : { query } }, { $group : { _groupby : { month: { $month: "$date" }, day: { $dayOfMonth: "$date" }, year: { $year: "$date" } }, new_field_1 : { $sum/$avg : "$key"} }} ]


' aggregation de donnees pour obtenir des tableaux synthetiques (ultra light pour l instant...)
    ' @author : gchevalley
    ' @version : 0.2
        'changes : correct bug if multiple aggregations for the same field : $avg age, $max age etc.. add $max and $min accumulator
    '
    ' @param oQuery : objet collection composee de :
                        '
                        ' d un objet '$match' faculatif contenant une requete de selection. Si cet element est omis, tous les documents de la collection sont utilises pour effectuer l aggregation
                        '
                        ' d un objet '$group' contenant :
                            '
                            ' un objet "_id" compose des entetes libres pointant sur les champs de texte que l on souhaite aggeges precede d un signe dollar (GROUP BY country, currency, sector etc.)
                                ' ex :
                                ' _id' : {'custom_header_group_by_1' : '$field_group_by_1', 'custom_header_group_by_2' : '$field_group_by_2', ..., 'custom_header_group_by_n' : '$field_group_by_n'}
                            '
                            ' de liste d objets numeriques devant etre accumules : custom_header_1 : {'$accumulator' : '$field'}
                                ' ex:
                                ' custom_header_accumlator_1' : {'$count'|'$sum'|'$max'|'$min'|'$avg' : '$field_to_accumulate_1'}, 'custom_header_accumlator_2' : {'$count'|'$sum'|'$max'|'$min'|'$avg' : '$field_to_accumulate_2'}, ..., 'custom_header_accumlator_m' : {'$count'|'$sum'|'$max'|'$min'|'$avg' : '$field_to_accumulate_m'}
                        
                        
            ' exemple complet d une structure de requete :
            '[
                '{$match' : {<query_de_selection>}},
                '
                '{$group' : {
                        '_id' : {
                            'custom_header_group_by_1' : '$field_group_by_1',
                            'custom_header_group_by_2' : '$field_group_by_2',
                            '...
                            'custom_header_group_by_n' : '$field_group_by_n'
                        '},
                        '
                        'custom_header_accumlator_1' : {'$count'|'$sum'|'$max'|'$min'|'$avg' : '$field_to_accumulate_1'},
                        'custom_header_accumlator_2' : {'$count'|'$sum'|'$max'|'$min'|'$avg' : '$field_to_accumulate_2'},
                        '...,
                        'custom_header_accumlator_m' : {'$count'|'$sum'|'$max'|'$min'|'$avg' : '$field_to_accumulate_m'}
                '}}
            ']
                    
    '
    ' @return : objet cls_NoSQL_QueryResult contenant un document par combinaison possible par rapport aux champs que l on souhaite aggreger (defini a travers l objet $group._id)
Public Function aggregate(oQuery) As cls_NoSQL_QueryResult

Set aggregate = New cls_NoSQL_QueryResult
Set aggregate.ref_collection = Me

Dim oJSON As New jsonlib

Dim oQueryCol As collection 'car ordre compte ici !

Set oQueryCol = transform_query_into_object(oQuery)


'y a - t - il une query $match ?
Dim oResult As cls_NoSQL_QueryResult
Dim found_match_query As Boolean
found_match_query = False
Dim tmp_dic As Scripting.Dictionary
For Each tmp_dic In oQueryCol
    If tmp_dic.Exists("$match") Then
        found_match_query = True
        Set oResult = Me.find(tmp_dic.Item("$match"))
    End If
Next

If found_match_query = False Then
    'take all
    Set oResult = New cls_NoSQL_QueryResult
    
    Dim tmp_oid As Variant
    For Each tmp_oid In Me.documents
        oResult.append_document Me.documents(tmp_oid)
    Next
End If



Dim dic_internal_operator As New Scripting.Dictionary
    dic_internal_operator.Add "$multiply", ""
    dic_internal_operator.Add "$avg", ""
    dic_internal_operator.Add "$sum", ""
    dic_internal_operator.Add "$count", ""


Dim found_group_query As Boolean
found_group_query = False

Dim oQueryGroup As Scripting.Dictionary


aggregate.query = oJSON.toString(oQueryCol)

For Each tmp_dic In oQueryCol
    If tmp_dic.Exists("$group") Then
        found_group_query = True
        Set oQueryGroup = tmp_dic.Item("$group")
        
        'extraction des colonnes group by
        Dim dic_group_by_id_field As Scripting.Dictionary
        Set dic_group_by_id_field = oQueryGroup.Item("_id")
        
        
        Dim dic_new_calc_field As New Scripting.Dictionary
        
        Dim tmp_field As Variant
        For Each tmp_field In oQueryGroup.keys
            
            If tmp_field <> "_id" Then
                dic_new_calc_field.Add tmp_field, oQueryGroup.Item(tmp_field)
            End If
            
        Next
        
        
        
        'passe en revue les documents et isole ceux qui dispose de toutes les colonnes
        Dim dic_doc_final_to_aggregate As New Scripting.Dictionary
        Dim count_doc_to_aggregate As Long
        count_doc_to_aggregate = 1
        Dim tmp_doc_to_check_necessary_fields As cls_NoSQL_Document
        For Each tmp_oid In oResult.documents.keys
            
            Dim take_doc As Boolean
            take_doc = True
            
            Set tmp_doc_to_check_necessary_fields = oResult.documents.Item(tmp_oid)
            
            For Each tmp_field In dic_group_by_id_field.keys
                If tmp_doc_to_check_necessary_fields.check_if_key_exits(Replace(dic_group_by_id_field.Item(tmp_field), "$", "")) = False Then
                    take_doc = False
                End If
            Next
            
            
            Dim dic_accumulator_and_field As Scripting.Dictionary
            
            For Each tmp_field In dic_new_calc_field.keys
                
                Set dic_accumulator_and_field = dic_new_calc_field.Item(tmp_field)
                
                'Debug.Print Replace(dic_accumulator_and_field.Item(dic_accumulator_and_field.keys(0)), "$", "")
                
                If tmp_doc_to_check_necessary_fields.check_if_key_exits(Replace(dic_accumulator_and_field.Item(dic_accumulator_and_field.keys(0)), "$", "")) = False Then
                    take_doc = False
                Else
                    
                    If IsNumeric(tmp_doc_to_check_necessary_fields.get_value_from_key(Replace(dic_accumulator_and_field.Item(dic_accumulator_and_field.keys(0)), "$", ""))) Then
                        
                    Else
                        take_doc = False
                    End If
                
                End If
            Next
            
            
            If take_doc Then
                
                'creation du nouveau document
                Dim dic_new_doc As Scripting.Dictionary
                Set dic_new_doc = New Scripting.Dictionary
                
                For Each tmp_field In dic_group_by_id_field.keys
                    dic_new_doc.Add Replace(dic_group_by_id_field.Item(tmp_field), "$", ""), tmp_doc_to_check_necessary_fields.get_value_from_key(Replace(dic_group_by_id_field.Item(tmp_field), "$", ""))
                Next
                
                
                For Each tmp_field In dic_new_calc_field.keys
                    Set dic_accumulator_and_field = dic_new_calc_field.Item(tmp_field)
                    
                    If dic_new_doc.Exists(Replace(dic_accumulator_and_field.Item(dic_accumulator_and_field.keys(0)), "$", "")) = False Then
                        dic_new_doc.Add Replace(dic_accumulator_and_field.Item(dic_accumulator_and_field.keys(0)), "$", ""), tmp_doc_to_check_necessary_fields.get_value_from_key(Replace(dic_accumulator_and_field.Item(dic_accumulator_and_field.keys(0)), "$", ""))
                    End If
                Next
                
                
                'Debug.Print dic_new_doc.Item("name")
                
                
                dic_doc_final_to_aggregate.Add CStr(count_doc_to_aggregate), dic_new_doc
                count_doc_to_aggregate = count_doc_to_aggregate + 1
            End If
            
        Next 'boucle sur oResult.documents
        
        
        'creation des distinct _id
        Dim dic_distinct_keys As New Scripting.Dictionary
        
        For Each tmp_oid In dic_doc_final_to_aggregate.keys
            
            Set dic_new_doc = dic_doc_final_to_aggregate.Item(tmp_oid)
            
            'construct distinct key
            Dim tmp_distinct_key As Variant
            tmp_distinct_key = ""
            
            Dim dic_key_with_value As Scripting.Dictionary
            Set dic_key_with_value = New Scripting.Dictionary
            For Each tmp_field In dic_group_by_id_field.keys
                'Debug.Print "key: " & Replace(dic_group_by_id_field.Item(tmp_field), "$", "") & " and value: " & dic_new_doc.Item(Replace(dic_group_by_id_field.Item(tmp_field), "$", ""))
                dic_key_with_value.Add Replace(dic_group_by_id_field.Item(tmp_field), "$", ""), dic_new_doc.Item(Replace(dic_group_by_id_field.Item(tmp_field), "$", ""))
                tmp_distinct_key = tmp_distinct_key & "*#*#*" & dic_new_doc.Item(Replace(dic_group_by_id_field.Item(tmp_field), "$", ""))
            Next
            
            If dic_distinct_keys.Exists(tmp_distinct_key) = False Then
                dic_distinct_keys.Add tmp_distinct_key, dic_key_with_value
            End If
            
        Next
        
        
        
        
        
        'passe en revue les distinct key
        Dim dic_doc_aggregate As New Scripting.Dictionary 'just before final with $sum and $count
        
        'passe en revue les keys
        Dim dic_tmp_doc As Scripting.Dictionary
        Dim dic_calcs As New Scripting.Dictionary
        For Each tmp_distinct_key In dic_distinct_keys.keys
            
            Set dic_key_with_value = dic_distinct_keys.Item(tmp_distinct_key)
            
            'passe en revue les documents
            Dim found_all_keys_values As Boolean
            found_all_keys_values = True
            For Each tmp_oid In dic_doc_final_to_aggregate.keys
                
                Set dic_tmp_doc = dic_doc_final_to_aggregate.Item(tmp_oid)
                
                found_all_keys_values = True
                For Each tmp_field In dic_key_with_value
                    
                    If dic_tmp_doc.Item(tmp_field) <> dic_key_with_value.Item(tmp_field) Then
                        found_all_keys_values = False
                    End If
                    
                Next
                
                
                Dim oColStatsValue As collection
                
                If found_all_keys_values Then
                    
                    Set dic_new_doc = New Scripting.Dictionary
                    
                    
                    If dic_doc_aggregate.Exists(tmp_distinct_key) Then
                        'ajoute content
                        Set dic_new_doc = dic_doc_aggregate.Item(tmp_distinct_key)
                        
                        For Each tmp_field In dic_new_doc.keys
                            If tmp_field <> "_id" Then
                                Set dic_calcs = dic_new_doc.Item(tmp_field)
                                
                                dic_calcs.Item("$sum") = dic_calcs.Item("$sum") + dic_tmp_doc.Item(tmp_field)
                                dic_calcs.Item("$count") = dic_calcs.Item("$count") + 1
                                
                                If dic_tmp_doc.Item(tmp_field) > dic_calcs.Item("$max") Then
                                    dic_calcs.Item("$max") = dic_tmp_doc.Item(tmp_field)
                                End If
                                
                                If dic_tmp_doc.Item(tmp_field) < dic_calcs.Item("$min") Then
                                    dic_calcs.Item("$min") = dic_tmp_doc.Item(tmp_field)
                                End If
                                
                                Set oColStatsValue = dic_calcs.Item("$values")
                                    oColStatsValue.Add dic_tmp_doc.Item(tmp_field)
                                    Set dic_calcs.Item("$values") = oColStatsValue
                                
                            End If
                        Next
                        
                    Else
                        
                        'creation de la zone key
                        dic_new_doc.Add "_id", dic_key_with_value
                        
                        For Each tmp_field In dic_new_calc_field.keys
                            
                            Set dic_accumulator_and_field = dic_new_calc_field.Item(tmp_field)
                            
                            'Debug.Print Replace(dic_accumulator_and_field.Item(dic_accumulator_and_field.keys(0)), "$", "")
                            
                            If dic_new_doc.Exists(Replace(dic_accumulator_and_field.Item(dic_accumulator_and_field.keys(0)), "$", "")) = False Then
                                Set dic_calcs = New Scripting.Dictionary
                                    
                                    'Debug.Print dic_tmp_doc.Item(Replace(dic_accumulator_and_field.Item(dic_accumulator_and_field.keys(0)), "$", ""))
                                
                                
                                Set oColStatsValue = New collection
                                    oColStatsValue.Add dic_tmp_doc.Item(Replace(dic_accumulator_and_field.Item(dic_accumulator_and_field.keys(0)), "$", ""))
                                
                                dic_calcs.Add "$min", dic_tmp_doc.Item(Replace(dic_accumulator_and_field.Item(dic_accumulator_and_field.keys(0)), "$", ""))
                                dic_calcs.Add "$max", dic_tmp_doc.Item(Replace(dic_accumulator_and_field.Item(dic_accumulator_and_field.keys(0)), "$", ""))
                                dic_calcs.Add "$sum", dic_tmp_doc.Item(Replace(dic_accumulator_and_field.Item(dic_accumulator_and_field.keys(0)), "$", ""))
                                dic_calcs.Add "$count", 1
                                dic_calcs.Add "$values", oColStatsValue
                                
                                dic_new_doc.Add Replace(dic_accumulator_and_field.Item(dic_accumulator_and_field.keys(0)), "$", ""), dic_calcs
                                
                            End If
                            
                        Next
                        
                        
                        dic_doc_aggregate.Add tmp_distinct_key, dic_new_doc
                        
                    End If
                End If
                
                
            Next
        Next
        
        
        
        Dim dic_final_output As New Scripting.Dictionary
        Dim dic_final_output_row As Scripting.Dictionary
        
        For Each tmp_distinct_key In dic_doc_aggregate
            
            Set dic_new_doc = dic_doc_aggregate.Item(tmp_distinct_key)
            
            Set dic_final_output_row = New Scripting.Dictionary
            
            dic_final_output_row.Add "_id", dic_doc_aggregate.Item(tmp_distinct_key).Item("_id")
            
            For Each tmp_field In dic_new_calc_field.keys
                
                Set dic_accumulator_and_field = dic_new_calc_field.Item(tmp_field)
                Set dic_calcs = dic_new_doc.Item(Replace(dic_accumulator_and_field.Item(dic_accumulator_and_field.keys(0)), "$", ""))
                
                Dim tmp_finale_value As Variant
                tmp_finale_value = Null
                
                Select Case dic_accumulator_and_field.keys(0)
                    Case "$avg"
                        If dic_calcs.Item("$count") > 0 Then
                            tmp_finale_value = dic_calcs.Item("$sum") / dic_calcs.Item("$count")
                        Else
                            tmp_finale_value = Null
                        End If
                    Case "$sum"
                        tmp_finale_value = dic_calcs.Item("$sum")
                    Case "$count"
                        tmp_finale_value = dic_calcs.Item("$count")
                    Case "$max"
                        tmp_finale_value = dic_calcs.Item("$max")
                    Case "$min"
                        tmp_finale_value = dic_calcs.Item("$min")
                    Case Else
                        
                End Select
                
                dic_final_output_row.Add tmp_field, tmp_finale_value
            Next
            
            Dim oDoc As cls_NoSQL_Document
            Set oDoc = New cls_NoSQL_Document
                oDoc.load_data dic_final_output_row
            
            aggregate.append_document oDoc
            
            'dic_final_output.Add tmp_distinct_key, dic_final_output_row 'deprecie, utilise directement la fonction aggregate elle meme qui est un objet cls_NoSQL_QueryResult
            
        Next
        
    Else
        
        Debug.Print "WARNING cls_NoSQL_Collection::aggregate object $group missing in query: " & oJSON.toString(oQueryCol) & ". Nothing to do -> quit."
        
    End If
Next


Debug.Print "INFO cls_NoSQL_Collection::aggregate query: " & oJSON.toString(oQueryCol) & " returns: " & aggregate.count & " result(s)"


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
