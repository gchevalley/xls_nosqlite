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
    Debug.Print "ERROR cls_NoSQL_Document::load_data no id no idea about the datatype"
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


Public Function edit_pair_key_value(ByVal key As String, oNewValue)

Dim oJSON As New jsonlib

If InStr(key, ".") = 0 Then

    If Me.representation_dictionary.Exists(key) Then
        
        'object ou simple variable
        If VarType(oNewValue) = vbObject Then
            Set Me.representation_dictionary.Item(key) = oNewValue
        Else
            Me.representation_dictionary.Item(key) = oNewValue
        End If
    Else
        Me.representation_dictionary.Add key, oNewValue
    End If

Else
    
    Dim key_last_level As String
    key_last_level = StrReverse(Left(StrReverse(key), InStr(StrReverse(key), ".") - 1))
    
    Dim col_objects_to_edit As Collection
    
    'comme la cle peut etre manquante, demande one level up
    Set col_objects_to_edit = Me.get_collection_sub_objects(Left(key, Len(key) - Len(key_last_level) - 1))
    
    If col_objects_to_edit.count > 0 Then
        
        'Debug.Print oJSON.toString(Me.representation_dictionary)
        
        Dim tmp_bloc_dic As Scripting.Dictionary
        For Each tmp_bloc_dic In col_objects_to_edit
            
            If tmp_bloc_dic.Exists(key_last_level) Then
                
                'object ou simple variable
                If VarType(oNewValue) = vbObject Then
                    Set tmp_bloc_dic.Item(key_last_level) = oNewValue
                Else
                    tmp_bloc_dic.Item(key_last_level) = oNewValue
                End If
                
            Else
                tmp_bloc_dic.Add key_last_level, oNewValue
            End If
            
            'Debug.Print oJSON.toString(tmp_bloc_dic)
        Next
        
        
        
        'Debug.Print oJSON.toString(Me.representation_dictionary)
        
    End If
    
End If

sync_update_represenation_json_from_representation_dictionary

End Function


Public Function get_collection_sub_objects(ByVal key As String) As Collection

'last level should be a dictionary object ! (not an array/collection)

Set get_collection_sub_objects = New Collection

Dim k As Long

Dim oReg As New VBScript_RegExp_55.RegExp
Dim match As VBScript_RegExp_55.match
Dim matches As VBScript_RegExp_55.MatchCollection
    oReg.Global = True


'embedded field ?
Dim array_CheckLevelConstraint As New Collection, array_CheckSubLevelConstraint As Collection
Dim dic_CheckLevelConstraint As Scripting.Dictionary ', dic_current_level As Scripting.Dictionary, dic_sub_level As Scripting.Dictionary

Set dic_CheckLevelConstraint = Me.representation_dictionary 'top level must be a dictionary

array_CheckLevelConstraint.Add dic_CheckLevelConstraint 'package tout ca dans une collection


If InStr(key, ".") <> 0 Then
    
    oReg.Pattern = "[\w]+"
    Set matches = oReg.Execute(key)
    
    k = 1
    For Each match In matches
        
        Set array_CheckSubLevelConstraint = New Collection
        
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
                    ElseIf TypeOf tmp_o.Item(match.value) Is Collection Then
                        'Set array_CheckSubLevelConstraint = tmp_o.Item(match.value) 'array d objects
                        Dim tmp_col As Collection
                        Set tmp_col = tmp_o.Item(match.value)
                        
                        Dim tmp_dic_in_collection 'As Scripting.Dictionary
                        For Each tmp_dic_in_collection In tmp_col
                            If VarType(tmp_dic_in_collection) = vbObject Then
                                array_CheckSubLevelConstraint.Add tmp_dic_in_collection
                            Else
'                                Debug.Print "WARNING cls_NoSQL_Document::get_collection_sub_objects query on scalar variable instead of object (array components). Should use query operator $in. Now cast scalar into compatible object"
'                                Dim tmp_dic_cast_scalar As Scripting.Dictionary
'                                Set tmp_dic_cast_scalar = New Scripting.Dictionary
'                                tmp_dic_cast_scalar.Add match.value, tmp_dic_in_collection
'
'                                array_CheckSubLevelConstraint.Add tmp_dic_cast_scalar
                                
                                
                                'last level ?
                                If k = matches.count Then
                                    'retourne full collection (=array)
                                    Dim tmp_array_collection As New Collection
                                    
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
