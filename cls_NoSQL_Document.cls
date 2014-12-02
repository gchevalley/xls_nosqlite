VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_NoSQL_Document"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public oid As String
'Public representation_json As String
'Public representation_dictionary As Scripting.Dictionary
Private prepresentation_json As String
Private prepresentation_dictionary As Scripting.Dictionary

Public ref_collection As cls_NoSQL_Collection

'Public Event needUpdateRepresenationJSON(ByRef Cancel As Boolean)


Private Sub Class_Initialize()

'Set ref_collection = New cls_NoSQL_Collection 'pour rester a nothing

End Sub



Public Property Get representation_json() As String

representation_json = prepresentation_json

End Property


Public Property Let representation_json(value As String)

prepresentation_json = value

sync_update_representation_dictionary_from_represenation_json

End Property


Public Property Get representation_dictionary() As Scripting.Dictionary

Set representation_dictionary = prepresentation_dictionary

End Property


Public Property Let representation_dictionary(value As Scripting.Dictionary)

Set prepresentation_dictionary = value

sync_update_represenation_json_from_representation_dictionary

End Property


'first level must absolutely be an object (dictionary)
Public Function load_data(o)

Dim oJSON As New jsonlib

If VarType(o) = vbObject Then
    If TypeOf o Is Scripting.Dictionary Then
        Me.representation_dictionary = o
        
'        If o.Exists("_id") Then
'            'Debug.Print "_id from cls_NoSQL_Document::load_data " & o.Item("_id") 'doit etre en dico, cette ligne doit planter
'            Debug.Print oJSON.toString(o)
'        End If
        
        
        'Me.representation_json = oJSON.toString(o)
    ElseIf TypeOf o Is cls_NoSQL_Document Then
        Me.representation_dictionary = o.representation_dictionary
        'Me.representation_json = o.representation_json
    Else
        Debug.Print "ERROR cls_NoSQL_Document::load_data no id about the kind of object"
    End If
ElseIf VarType(o) = vbString Then
    Me.representation_json = o
    'Me.representation_dictionary = oJSON.parse(CStr(o))
Else
    Debug.Print "ERROR cls_NoSQL_Document::load_data no idea about the datatype => should a dictionary on top level"
End If

'existe t il deja une cle?
If Me.representation_dictionary.Exists("_id") Then
    oid = oJSON.toString(Me.representation_dictionary.Item("_id"))
    'oid = Me.representation_dictionary.Item("_id")
End If

End Function

Public Function link_with_collection(oCollection As cls_NoSQL_Collection)

Set ref_collection = oCollection

End Function


Public Function check_if_key_exits(ByVal key As String) As Boolean

Dim oJSON As New jsonlib

If InStr(key, ".") = 0 Then
    check_if_key_exits = Me.representation_dictionary.Exists(key)
Else
    
    Dim key_last_level As String
    key_last_level = StrReverse(Left(StrReverse(key), InStr(StrReverse(key), ".") - 1))
    
    Dim col_objects_to_edit As collection
    
    Set col_objects_to_edit = Me.get_collection_sub_objects(key)
    
    If col_objects_to_edit.count > 0 Then
        
        Dim tmp_bloc_dic As Scripting.Dictionary
        
        Dim check_object_type As Object
        
        For Each check_object_type In col_objects_to_edit
            
            If TypeOf check_object_type Is Scripting.Dictionary Then
                
                Set tmp_bloc_dic = check_object_type
                
                If tmp_bloc_dic.Exists(key_last_level) Then
                    check_if_key_exits = True
                End If
            
            ElseIf TypeOf check_object_type Is collection Then
                
                'remote un level up
                Dim col_objects_to_edit_from_collection As collection
                Set col_objects_to_edit_from_collection = Me.get_collection_sub_objects(Left(key, Len(key) - Len(key_last_level) - 1))
                
                'effectue le changement
                For Each tmp_bloc_dic In col_objects_to_edit_from_collection 'boucle sur les objets de l'array
                    If tmp_bloc_dic.Exists(key_last_level) Then
                        check_if_key_exits = True
                    End If
                Next
                
            End If
        Next
            
        
    End If
    
End If

End Function


Public Function get_value_from_key(ByVal key As String) As Variant

get_value_from_key = Empty

Dim oJSON As New jsonlib

If InStr(key, ".") = 0 Then
    If Me.representation_dictionary.Exists(key) Then
        get_value_from_key = Me.representation_dictionary.Item(key)
    End If
Else
    
    Dim key_last_level As String
    key_last_level = StrReverse(Left(StrReverse(key), InStr(StrReverse(key), ".") - 1))
    
    Dim col_objects_to_edit As collection
    
    Set col_objects_to_edit = Me.get_collection_sub_objects(key)
    
    If col_objects_to_edit.count > 0 Then
        
        Dim tmp_bloc_dic As Scripting.Dictionary
        
        Dim check_object_type As Object
        
        For Each check_object_type In col_objects_to_edit
            
            If TypeOf check_object_type Is Scripting.Dictionary Then
                
                Set tmp_bloc_dic = check_object_type
                
                If tmp_bloc_dic.Exists(key_last_level) Then
                    get_value_from_key = tmp_bloc_dic.Item(key_last_level)
                End If
            
            ElseIf TypeOf check_object_type Is collection Then
                
                ' glurb....
                
            End If
        Next
            
        
    End If
    
End If



End Function


Public Function edit_pair_key_value(ByVal key As String, oNewValue)

Dim oJSON As New jsonlib

If InStr(key, ".") = 0 Then

    If Me.representation_dictionary.Exists(key) Then
        
        'object ou simple variable
        If VarType(oNewValue) = vbObject Then
            Set Me.representation_dictionary.Item(key) = oNewValue
            Debug.Print oJSON.toString(Me.representation_dictionary)
        Else
            Me.representation_dictionary.Item(key) = oNewValue
            'Debug.Print oJSON.toString(Me.representation_dictionary)
        End If
    Else
        Me.representation_dictionary.Add key, oNewValue
    End If

Else
    
    Dim key_last_level As String
    key_last_level = StrReverse(Left(StrReverse(key), InStr(StrReverse(key), ".") - 1))
    
    Dim col_objects_to_edit As collection
    
    'comme la cle peut etre manquante, retourne one level up si dispo uniquement
    Set col_objects_to_edit = Me.get_collection_sub_objects(key)
    
    If col_objects_to_edit.count > 0 Then
        
        'Debug.Print oJSON.toString(Me.representation_dictionary)
        
        Dim tmp_bloc_dic As Scripting.Dictionary
        
        Dim check_object_type As Object
        
        For Each check_object_type In col_objects_to_edit
            
            If TypeOf check_object_type Is Scripting.Dictionary Then 'le sous objet est un dico
                
                Set tmp_bloc_dic = check_object_type
                
                If tmp_bloc_dic.Exists(key_last_level) Then
                    
                    Debug.Print "cls_NoSQL_Document::edit_pair_key_value on dictionary object"
                    
                    'object ou simple variable
                    If VarType(oNewValue) = vbObject Then
                        Set tmp_bloc_dic.Item(key_last_level) = oNewValue
                    Else
                        tmp_bloc_dic.Item(key_last_level) = oNewValue
                    End If
                    
                Else
                    tmp_bloc_dic.Add key_last_level, oNewValue
                End If
            
            ElseIf TypeOf check_object_type Is collection Then 'le sous-objet est un array constitue d objets
                
                'remote un level up
                Dim col_objects_to_edit_from_collection As collection
                Set col_objects_to_edit_from_collection = Me.get_collection_sub_objects(Left(key, Len(key) - Len(key_last_level) - 1))
                
                'effectue le changement
                For Each tmp_bloc_dic In col_objects_to_edit_from_collection 'boucle sur les objets de l'array
                    'object ou simple variable
                    
                    If tmp_bloc_dic.Exists(key_last_level) Then
                        
                        Debug.Print "cls_NoSQL_Document::edit_pair_key_value on dictionary object in an array"
                        
                        If VarType(oNewValue) = vbObject Then
                            Set tmp_bloc_dic.Item(key_last_level) = oNewValue
                        Else
                            tmp_bloc_dic.Item(key_last_level) = oNewValue
                        End If
                    Else
                        tmp_bloc_dic.Add key_last_level, oNewValue
                    End If
                Next
                
            End If
            
            'Debug.Print oJSON.toString(tmp_bloc_dic)
        Next
        
        
        
        'Debug.Print oJSON.toString(Me.representation_dictionary)
        
    End If
    
End If

sync_update_represenation_json_from_representation_dictionary

End Function


Public Function remove_key(ByVal key As String)

Dim oJSON As New jsonlib

If InStr(key, ".") = 0 Then
    
    If Me.representation_dictionary.Exists(key) Then
        Debug.Print "INFO cls_NoSQL_Document::remove_key key:" & key & " in " & Me.representation_json
        Me.representation_dictionary.remove key
    Else
        Debug.Print "INFO cls_NoSQL_Document::remove_key key: " & key & " not available in " & Me.representation_json
    End If
    
Else
    
    Dim key_last_level As String
    key_last_level = StrReverse(Left(StrReverse(key), InStr(StrReverse(key), ".") - 1))
    
    Dim col_objects_to_edit As collection
    Set col_objects_to_edit = Me.get_collection_sub_objects(key)
    
    If col_objects_to_edit.count > 0 Then
        
        'destruction de la cle sur les sous-objets qui repercute les modif sur l objet global !
        Dim tmp_dic_unset_top_level As Scripting.Dictionary, tmp_dic_sub_level As Scripting.Dictionary
        For Each tmp_dic_unset_top_level In col_objects_to_edit
            
            If tmp_dic_unset_top_level.Exists(key_last_level) Then
                
                Debug.Print "INFO cls_NoSQL_Document::remove_key key:" & key & " in " & Me.representation_json
                tmp_dic_unset_top_level.remove key_last_level
                
                'controle du fait qu appliquer une modif sur un sous objet applique bien sur le tout
                'Debug.Print oJSON.toString(Me.representation_dictionary)
            Else
                Debug.Print "INFO cls_NoSQL_Document::remove_key key: " & key & " not available in " & Me.representation_json
            End If
            
        Next
    Else
        Debug.Print "INFO cls_NoSQL_Document::remove_key nothing to do. 0 sub-object"
    End If
End If


sync_update_represenation_json_from_representation_dictionary

End Function


Public Function rename_key(ByVal old_key As String, ByVal new_key As String)

Dim oJSON As New jsonlib

'traitement de new key
If InStr(new_key, ".") Then
    new_key = StrReverse(Left(StrReverse(new_key), InStr(StrReverse(new_key), ".") - 1))
End If


Dim tmp_scalar As Variant
Dim tmp_o As Object

If InStr(old_key, ".") = 0 Then
    
    If Me.representation_dictionary.Exists(old_key) Then
        
        
        
        If VarType(Me.representation_dictionary.Item(old_key)) = vbObject Then
            Set tmp_o = Me.representation_dictionary.Item(old_key)
            
            Me.representation_dictionary.remove old_key
            Me.representation_dictionary.Add new_key, tmp_o
        Else
            tmp_scalar = Me.representation_dictionary.Item(old_key)
            
            Me.representation_dictionary.remove old_key
            Me.representation_dictionary.Add new_key, tmp_scalar
        End If
        
    Else
        Debug.Print "INFO cls_NoSQL_Document::rename_key key: " & old_key & " not available in " & Me.representation_json
    End If
    
Else
    
    Dim key_last_level As String
    key_last_level = StrReverse(Left(StrReverse(old_key), InStr(StrReverse(old_key), ".") - 1))
    
    Dim col_objects_to_edit As collection
    Set col_objects_to_edit = Me.get_collection_sub_objects(old_key)
    
    If col_objects_to_edit.count > 0 Then
        
        'destruction de la cle sur les sous-objets qui repercute les modif sur l objet global !
        Dim tmp_dic_rename_top_level As Scripting.Dictionary, tmp_dic_sub_level As Scripting.Dictionary
        For Each tmp_dic_rename_top_level In col_objects_to_edit
            
            If tmp_dic_rename_top_level.Exists(key_last_level) Then
                
                Debug.Print "INFO before cls_NoSQL_Document::rename_key key:" & old_key & " in " & Me.representation_json
                
                If VarType(tmp_dic_rename_top_level.Item(key_last_level)) = vbObject Then
                    Set tmp_o = tmp_dic_rename_top_level.Item(key_last_level)
                    
                    tmp_dic_rename_top_level.remove key_last_level
                    tmp_dic_rename_top_level.Add new_key, tmp_o
                Else
                    tmp_scalar = tmp_dic_rename_top_level.Item(key_last_level)
                    
                    tmp_dic_rename_top_level.remove key_last_level
                    tmp_dic_rename_top_level.Add new_key, tmp_scalar
                End If
                
                'controle du fait qu appliquer une modif sur un sous objet applique bien sur le tout
                Debug.Print oJSON.toString(Me.representation_dictionary)
            Else
                Debug.Print "INFO cls_NoSQL_Document::rename_key key: " & old_key & " not available in " & Me.representation_json
            End If
            
        Next
    Else
        Debug.Print "INFO cls_NoSQL_Document::rename_key nothing to do. 0 sub-object"
    End If
End If


sync_update_represenation_json_from_representation_dictionary

End Function


'side = 'min' / 'max'
Public Function min_max_value_for_one_key(ByVal side As String, ByVal key As String, ByVal value As Double) As Boolean

min_max_value_for_one_key = False

Dim oJSON As New jsonlib

If InStr(key, ".") = 0 Then

    If Me.representation_dictionary.Exists(key) Then
        

        If IsNumeric(Me.representation_dictionary.Item(key)) Then
            'Debug.Print Evaluate(value & side & Me.representation_dictionary.Item(key))
            
            If Evaluate(value & side & Me.representation_dictionary.Item(key)) Then
                Debug.Print "INFO cls_NoSQL_Document::min_max_value_for_one_key update key: " & key & " with new numerical value: " & value
                Me.representation_dictionary.Item(key) = value
                min_max_value_for_one_key = True
            Else
                Debug.Print "INFO cls_NoSQL_Document::min_max_value_for_one_key value: " & Me.representation_dictionary.Item(key) & " already out of limit: " & value
            End If
            
        Else
            Debug.Print "WARNING cls_NoSQL_Document::min_max_value_for_one_key value for key: " & key & " is not numerical"
        End If
    
    Else
        
        Debug.Print "INFO cls_NoSQL_Document::min_max_value_for_one_key missing key: " & key & ". Add new one"
        Me.representation_dictionary.Add key, value
        min_max_value_for_one_key = True
        
    End If


Else
    
    Dim key_last_level As String
    key_last_level = StrReverse(Left(StrReverse(key), InStr(StrReverse(key), ".") - 1))
    
    Dim col_objects_to_edit As collection
    
    'comme la cle peut etre manquante, retourne one level up si dispo uniquement
    Set col_objects_to_edit = Me.get_collection_sub_objects(key)
    
    If col_objects_to_edit.count > 0 Then
        
        'Debug.Print oJSON.toString(Me.representation_dictionary)
        
        Dim tmp_bloc_dic As Scripting.Dictionary
        
        Dim check_object_type As Object
        
        For Each check_object_type In col_objects_to_edit
            
            If TypeOf check_object_type Is Scripting.Dictionary Then 'le sous objet est un dico
                
                Set tmp_bloc_dic = check_object_type
                
                If tmp_bloc_dic.Exists(key_last_level) Then
                    
                    If IsNumeric(tmp_bloc_dic.Item(key_last_level)) Then
                        
                        If Evaluate(value & side & tmp_bloc_dic.Item(key_last_level)) Then
                            Debug.Print "cls_NoSQL_Document::min_max_value_for_one_key edit dictionary object, key: " & key & " with new value: " & value
                            tmp_bloc_dic.Item(key_last_level) = value
                            min_max_value_for_one_key = True
                        Else
                            Debug.Print "INFO cls_NoSQL_Document::min_max_value_for_one_key sub-object value: " & tmp_bloc_dic.Item(key_last_level) & " already out of limit: " & value
                        End If
                    Else
                        Debug.Print "WARNING cls_NoSQL_Document::min_max_value_for_one_key value for sub-key: " & key & " is not numerical"
                    End If
                    
                Else
                    Debug.Print "INFO cls_NoSQL_Document::min_max_value_for_one_key missing key in sub-object: " & key & ". Add new one"
                    tmp_bloc_dic.Add key_last_level, value
                    min_max_value_for_one_key = True
                End If
            
            ElseIf TypeOf check_object_type Is collection Then 'le sous-objet est un array constitue d objets
                
                'remote un level up
                Dim col_objects_to_edit_from_collection As collection
                Set col_objects_to_edit_from_collection = Me.get_collection_sub_objects(Left(key, Len(key) - Len(key_last_level) - 1))
                
                'effectue le changement
                For Each tmp_bloc_dic In col_objects_to_edit_from_collection 'boucle sur les objets de l'array
                    
                    If tmp_bloc_dic.Exists(key_last_level) Then
                        
                        If IsNumeric(tmp_bloc_dic.Item(key_last_level)) Then
                            
                            If Evaluate(value & side & tmp_bloc_dic.Item(key_last_level)) Then
                                Debug.Print "cls_NoSQL_Document::min_max_value_for_one_key edit one of the dictionaies from array, key: " & key & " with new value: " & value
                                tmp_bloc_dic.Item(key_last_level) = value
                                min_max_value_for_one_key = True
                            Else
                                Debug.Print "INFO cls_NoSQL_Document::min_max_value_for_one_key sub-object from array value: " & tmp_bloc_dic.Item(key_last_level) & " already out of limit: " & value
                            End If
                        Else
                            Debug.Print "WARNING cls_NoSQL_Document::min_max_value_for_one_key value for sub-key from array: " & key & " is not numerical"
                        End If
                        
                    Else
                        Debug.Print "INFO cls_NoSQL_Document::min_max_value_for_one_key missing key in sub-object dictionary from array: " & key & ". Add new one"
                        tmp_bloc_dic.Add key_last_level, value
                        min_max_value_for_one_key = True
                    End If
                Next
                
            End If
            
            'Debug.Print oJSON.toString(tmp_bloc_dic)
        Next
        
        
        
        'Debug.Print oJSON.toString(Me.representation_dictionary)
        
    End If
    
End If

sync_update_represenation_json_from_representation_dictionary

End Function


Public Function dephasing_value_for_one_key(ByVal key As String, ByVal differential As Double) As Boolean

dephasing_value_for_one_key = False

Dim oJSON As New jsonlib

If InStr(key, ".") = 0 Then

    If Me.representation_dictionary.Exists(key) Then
        

        If IsNumeric(Me.representation_dictionary.Item(key)) Then
            
            Debug.Print "INFO cls_NoSQL_Document::dephasing_value_for_one_key update key: " & key & " with new numerical value: " & Me.representation_dictionary.Item(key) + differential & " old value: " & Me.representation_dictionary.Item(key)
            Me.representation_dictionary.Item(key) = Me.representation_dictionary.Item(key) + differential
            dephasing_value_for_one_key = True
            
        Else
            Debug.Print "WARNING cls_NoSQL_Document::dephasing_value_for_one_key value for key: " & key & " is not numerical"
        End If
    
    Else
        
        Debug.Print "WARNING cls_NoSQL_Document::dephasing_value_for_one_key missing key: " & key & ". Add new one"
        'Me.representation_dictionary.Add key, differential
        'dephasing_value_for_one_key = True
        
    End If


Else
    
    Dim key_last_level As String
    key_last_level = StrReverse(Left(StrReverse(key), InStr(StrReverse(key), ".") - 1))
    
    Dim col_objects_to_edit As collection
    
    'comme la cle peut etre manquante, retourne one level up si dispo uniquement
    Set col_objects_to_edit = Me.get_collection_sub_objects(key)
    
    If col_objects_to_edit.count > 0 Then
        
        'Debug.Print oJSON.toString(Me.representation_dictionary)
        
        Dim tmp_bloc_dic As Scripting.Dictionary
        
        Dim check_object_type As Object
        
        For Each check_object_type In col_objects_to_edit
            
            If TypeOf check_object_type Is Scripting.Dictionary Then 'le sous objet est un dico
                
                Set tmp_bloc_dic = check_object_type
                
                If tmp_bloc_dic.Exists(key_last_level) Then
                    
                    If IsNumeric(tmp_bloc_dic.Item(key_last_level)) Then
                        
                        Debug.Print "cls_NoSQL_Document::dephasing_value_for_one_key edit dictionary object, key: " & key & " with new value: " & tmp_bloc_dic.Item(key_last_level) + differential
                        tmp_bloc_dic.Item(key_last_level) = tmp_bloc_dic.Item(key_last_level) + differential
                        dephasing_value_for_one_key = True

                    Else
                        Debug.Print "WARNING cls_NoSQL_Document::dephasing_value_for_one_key value for sub-key: " & key & " is not numerical"
                    End If
                    
                Else
                    Debug.Print "WARNING cls_NoSQL_Document::dephasing_value_for_one_key missing key in sub-object: " & key
                    'tmp_bloc_dic.Add key_last_level, differential
                    'dephasing_value_for_one_key = True
                End If
            
            ElseIf TypeOf check_object_type Is collection Then 'le sous-objet est un array constitue d objets
                
                'remote un level up
                Dim col_objects_to_edit_from_collection As collection
                Set col_objects_to_edit_from_collection = Me.get_collection_sub_objects(Left(key, Len(key) - Len(key_last_level) - 1))
                
                'effectue le changement
                For Each tmp_bloc_dic In col_objects_to_edit_from_collection 'boucle sur les objets de l'array
                    
                    If tmp_bloc_dic.Exists(key_last_level) Then
                        
                        If IsNumeric(tmp_bloc_dic.Item(key_last_level)) Then
                            Debug.Print "cls_NoSQL_Document::dephasing_value_for_one_key edit one of the dictionaies from array, key: " & key & " with new value: " & tmp_bloc_dic.Item(key_last_level) + differential
                            tmp_bloc_dic.Item(key_last_level) = tmp_bloc_dic.Item(key_last_level) + differential
                            dephasing_value_for_one_key = True
                        Else
                            Debug.Print "WARNING cls_NoSQL_Document::dephasing_value_for_one_key value for sub-key from array: " & key & " is not numerical"
                        End If
                        
                    Else
                        Debug.Print "INFO cls_NoSQL_Document::dephasing_value_for_one_key missing key in sub-object dictionary from array: " & key
                        'tmp_bloc_dic.Add key_last_level, differential
                        'dephasing_value_for_one_key = True
                    End If
                Next
                
            End If
            
            'Debug.Print oJSON.toString(tmp_bloc_dic)
        Next
        
        
        
        'Debug.Print oJSON.toString(Me.representation_dictionary)
        
    End If
    
End If

sync_update_represenation_json_from_representation_dictionary

End Function


Public Function append_value_to_an_array(ByVal key As String, oNewValue) As Boolean

append_value_to_an_array = False

Dim col_bloc_to_edit As collection

Dim final_array_to_edit As collection

If InStr(key, ".") = 0 Then
    
    If Me.representation_dictionary.Exists(key) Then
        
        If TypeOf Me.representation_dictionary.Item(key) Is collection Then
            Debug.Print "INFO cls_NoSQL_Document::append_value_to_an_array to existing array with key: " & key
            Set final_array_to_edit = Me.representation_dictionary.Item(key)
            final_array_to_edit.Add oNewValue
            append_value_to_an_array = True
        Else
            Debug.Print "ERROR cls_NoSQL_Document::append_value_to_an_array key: " & key & " doesn't contain an array !"
        End If
    Else
        
        Debug.Print "INFO cls_NoSQL_Document::append_value_to_an_array key: " & key & " is missing. Will be added with an array"
        
        Set final_array_to_edit = New collection
        final_array_to_edit.Add oNewValue
        
        Me.representation_dictionary.Add key, final_array_to_edit
        append_value_to_an_array = True
    End If
    
Else
    
    Set col_bloc_to_edit = get_collection_sub_objects(key)
    
    Dim key_last_level As String
    key_last_level = StrReverse(Left(StrReverse(key), InStr(StrReverse(key), ".") - 1))
    
    
    If col_bloc_to_edit.count > 0 Then
        
        Dim tmp_obj_to_check_typ As Object
        
        For Each tmp_obj_to_check_typ In col_bloc_to_edit
            
            If TypeOf tmp_obj_to_check_typ Is collection Then
                
                Debug.Print "cls_NoSQL_Document::append_value_to_an_array edit sub-array object with key: " & key
                
                Set final_array_to_edit = tmp_obj_to_check_typ
                
                final_array_to_edit.Add oNewValue
                append_value_to_an_array = True
            Else
                Debug.Print "ERROR cls_NoSQL_Document::append_value_to_an_array sub-object key: " & key & " doesn't contain an array !"
            End If
            
        Next
    
    Else
        
        'key missing
        Dim dic_one_level_up As Scripting.Dictionary
        
        Dim col_bloc_one_level_up As collection
        Set col_bloc_one_level_up = get_collection_sub_objects(Left(key, Len(key) - Len(key_last_level) - 1))
        
        If col_bloc_one_level_up.count = 1 Then 'devrait etre un dico
            For Each tmp_obj_to_check_typ In col_bloc_one_level_up
                If TypeOf tmp_obj_to_check_typ Is Scripting.Dictionary Then
                    
                    Debug.Print "INFO cls_NoSQL_Document::append_value_to_an_array key: " & key & " is missing. Will be added with an array"
                    
                    Set final_array_to_edit = New collection
                    final_array_to_edit.Add oNewValue
        
                    Set dic_one_level_up = tmp_obj_to_check_typ
                    
                    dic_one_level_up.Add key_last_level, final_array_to_edit
                    
                    append_value_to_an_array = True
                    
                End If
            Next
        End If
        
    End If

End If

sync_update_represenation_json_from_representation_dictionary

End Function


'value -1 => first, value 1 => last
Public Function pop_element_from_an_array(ByVal key As String, ByVal value As Long) As Boolean

Dim index_to_remove As Long

pop_element_from_an_array = False

Dim col_bloc_to_edit As collection

Dim final_array_to_edit As collection

If InStr(key, ".") = 0 Then
    
    If Me.representation_dictionary.Exists(key) Then
        
        If TypeOf Me.representation_dictionary.Item(key) Is collection Then
            Debug.Print "INFO cls_NoSQL_Document::pop_element_from_an_array to existing array with key: " & key
            Set final_array_to_edit = Me.representation_dictionary.Item(key)
            
            If value < 0 Then
                index_to_remove = Abs(value) 'ex first
            Else
                index_to_remove = final_array_to_edit.count - (value - 1) 'ex last
            End If
            
            If index_to_remove > 0 And index_to_remove <= final_array_to_edit.count Then
                final_array_to_edit.remove index_to_remove
                pop_element_from_an_array = True
            Else
                Debug.Print "ERROR cls_NoSQL_Document::pop_element_from_an_array index to remove out of range on array"
            End If
            
        Else
            Debug.Print "ERROR cls_NoSQL_Document::pop_element_from_an_array key: " & key & " doesn't contain an array !"
        End If
    Else
        Debug.Print "INFO cls_NoSQL_Document::pop_element_from_an_array key: " & key & " is missing"
    End If
    
Else
    
    Set col_bloc_to_edit = get_collection_sub_objects(key)
    
    Dim key_last_level As String
    key_last_level = StrReverse(Left(StrReverse(key), InStr(StrReverse(key), ".") - 1))
    
    
    If col_bloc_to_edit.count > 0 Then
        
        Dim tmp_obj_to_check_typ As Object
        
        For Each tmp_obj_to_check_typ In col_bloc_to_edit
            
            If TypeOf tmp_obj_to_check_typ Is collection Then
                
                Debug.Print "cls_NoSQL_Document::pop_element_from_an_array edit sub-array object with key: " & key
                
                Set final_array_to_edit = tmp_obj_to_check_typ
                
                If value < 0 Then
                    index_to_remove = Abs(value) 'ex first
                Else
                    index_to_remove = final_array_to_edit.count - (value - 1) 'ex last
                End If
                
                If index_to_remove > 0 And index_to_remove <= final_array_to_edit.count Then
                    final_array_to_edit.remove index_to_remove
                    pop_element_from_an_array = True
                Else
                    Debug.Print "ERROR cls_NoSQL_Document::pop_element_from_an_array index to remove out of range on sub-array"
                End If
            Else
                Debug.Print "ERROR cls_NoSQL_Document::pop_element_from_an_array sub-object key: " & key & " doesn't contain an array !"
            End If
            
        Next
    
    Else
        
        Debug.Print "INFO cls_NoSQL_Document::pop_element_from_an_array sub-key: " & key & " is missing"
        
    End If

End If

sync_update_represenation_json_from_representation_dictionary

End Function


Public Function remove_element_from_an_array_based_on_content(ByVal key As String, ByVal query As Variant) As Boolean

remove_element_from_an_array_based_on_content = False
Dim count_element_to_remove As Long
Dim vec_index_element_to_remove() As Variant

Dim k As Long

Dim col_bloc_to_edit As collection

Dim final_array_to_edit As collection

If InStr(key, ".") = 0 Then
    
    If Me.representation_dictionary.Exists(key) Then
        
        If TypeOf Me.representation_dictionary.Item(key) Is collection Then
            Debug.Print "INFO cls_NoSQL_Document::remove_element_from_an_array_based_on_content to existing array with key: " & key
            Set final_array_to_edit = Me.representation_dictionary.Item(key)
            
            remove_element_from_an_array_based_on_content = True
            
        Else
            Debug.Print "ERROR cls_NoSQL_Document::remove_element_from_an_array_based_on_content key: " & key & " doesn't contain an array !"
        End If
    Else
        Debug.Print "INFO cls_NoSQL_Document::remove_element_from_an_array_based_on_content key: " & key & " is missing in one level up dictionary"
    End If
    
Else
    
    Set col_bloc_to_edit = get_collection_sub_objects(key)
    
    Dim key_last_level As String
    key_last_level = StrReverse(Left(StrReverse(key), InStr(StrReverse(key), ".") - 1))
    
    
    If col_bloc_to_edit.count > 0 Then
        
        Dim tmp_obj_to_check_typ As Object
        
        For Each tmp_obj_to_check_typ In col_bloc_to_edit
            
            If TypeOf tmp_obj_to_check_typ Is collection Then
                
                Debug.Print "cls_NoSQL_Document::remove_element_from_an_array_based_on_content edit sub-array object with key: " & key
                
                Set final_array_to_edit = tmp_obj_to_check_typ
                
                remove_element_from_an_array_based_on_content = True
            Else
                Debug.Print "ERROR cls_NoSQL_Document::remove_element_from_an_array_based_on_content sub-object key: " & key & " doesn't contain an array !"
            End If
            
        Next
    
    Else
        
        Debug.Print "INFO cls_NoSQL_Document::remove_element_from_an_array_based_on_content sub-key: " & key & " is missing"
        
    End If

End If



'passe en revue les entrees
If remove_element_from_an_array_based_on_content = True Then
    Dim tmp_o
    k = 1
    count_element_to_remove = 0
    For Each tmp_o In final_array_to_edit
        If VarType(tmp_o) <> vbObject Then  'scalar
            If tmp_o = query Then
                
                Debug.Print "INFO cls_NoSQL_Document::remove_element_from_an_array_based_on_content found element matching the query"
                
                ReDim Preserve vec_index_element_to_remove(count_element_to_remove)
                vec_index_element_to_remove(count_element_to_remove) = k
                count_element_to_remove = count_element_to_remove + 1
            End If
        End If
        
        k = k + 1
    Next
    
    
    If count_element_to_remove > 0 Then
        
        remove_element_from_an_array_based_on_content = True
        
        For k = UBound(vec_index_element_to_remove, 1) To 0 Step -1 'a parcourir a l'envers !
            final_array_to_edit.remove vec_index_element_to_remove(k)
        Next k
    
    Else
        
        remove_element_from_an_array_based_on_content = False
    
    End If
End If

sync_update_represenation_json_from_representation_dictionary

End Function


Public Function get_collection_sub_objects(ByVal key As String) As collection

'last level should be a dictionary object ! (not an array/collection)

Set get_collection_sub_objects = New collection

Dim k As Long

Dim oReg As New VBScript_RegExp_55.RegExp
Dim match As VBScript_RegExp_55.match
Dim matches As VBScript_RegExp_55.MatchCollection
    oReg.Global = True


'embedded field ?
Dim array_CheckLevelConstraint As New collection, array_CheckSubLevelConstraint As collection
Dim dic_CheckLevelConstraint As Scripting.Dictionary ', dic_current_level As Scripting.Dictionary, dic_sub_level As Scripting.Dictionary

Set dic_CheckLevelConstraint = Me.representation_dictionary 'top level must be a dictionary

array_CheckLevelConstraint.Add dic_CheckLevelConstraint 'package tout ca dans une collection


If InStr(key, ".") <> 0 Then
    
    oReg.Pattern = "[\w]+"
    Set matches = oReg.Execute(key)
    
    k = 1
    For Each match In matches
        
        Set array_CheckSubLevelConstraint = New collection
        
        'passe en revue les objects (dictionary) de la collection
        Dim tmp_o As Scripting.Dictionary
        
        Dim take_candidate As Boolean
        For Each tmp_o In array_CheckLevelConstraint
            
            take_candidate = True
            
            
            If tmp_o.Exists(match.value) Then 'ce level contient il la cle du sublevel suivant ?
                
            Else
                take_candidate = False
            End If
            
            
            
            If take_candidate = True Then
                
                If VarType(tmp_o.Item(match.value)) = vbObject Then
                    If TypeOf tmp_o.Item(match.value) Is Scripting.Dictionary Then
                        array_CheckSubLevelConstraint.Add tmp_o.Item(match.value) 'single object
                    ElseIf TypeOf tmp_o.Item(match.value) Is collection Then
                        Dim tmp_col As collection
                        Set tmp_col = tmp_o.Item(match.value)
                        
                        Dim tmp_dic_in_collection 'As Scripting.Dictionary
                        For Each tmp_dic_in_collection In tmp_col
                            If VarType(tmp_dic_in_collection) = vbObject Then
                                array_CheckSubLevelConstraint.Add tmp_dic_in_collection
                            Else
                                
                                'last level ?
                                If k = matches.count Then
                                    'retourne full collection (=array)
                                    Dim tmp_array_collection As New collection
                                    
                                    array_CheckSubLevelConstraint.Add tmp_col
                                    Set get_collection_sub_objects = array_CheckSubLevelConstraint
                                    Exit Function
                                End If
                                
                                
                            End If
                        Next
                        
                    End If
                Else
                    array_CheckSubLevelConstraint.Add tmp_o 'ne descend pas plus bas 'scalaire
                End If
                
            End If
        
        Next
        
        Set array_CheckLevelConstraint = array_CheckSubLevelConstraint 'next -> current
        k = k + 1
    Next
    
    Set get_collection_sub_objects = array_CheckLevelConstraint
    
Else 'normal condition top level dictionary
    get_collection_sub_objects.Add Me.representation_dictionary
End If



End Function


Public Function sync_update_represenation_json_from_representation_dictionary()

Dim oJSON As New jsonlib
'Me.representation_json = oJSON.toString(Me.representation_dictionary)
'Debug.Print oJSON.toString(Me.representation_dictionary)
'Dim o As Scripting.Dictionary
'Set o = oJSON.parse(oJSON.toString(Me.representation_dictionary))
prepresentation_json = oJSON.toString(Me.representation_dictionary)

End Function


Public Function sync_update_representation_dictionary_from_represenation_json()

Dim oJSON As New jsonlib
'Set Me.representation_dictionary = oJSON.parse(CStr(Me.representation_json))
Set prepresentation_dictionary = oJSON.parse(CStr(Me.representation_json))

End Function

Private Sub Class_Terminate()

'Me.representation_dictionary = Nothing
representation_json = ""

End Sub
