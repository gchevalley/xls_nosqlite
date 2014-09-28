VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_NoSQL_QueryResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ref_collection As cls_NoSQL_Collection
Public query As String
Public documents As Scripting.Dictionary
Public orders_keys As Variant
Public count As Long

Public Function sort(oQuery)

Dim k As Long, m As Long, n As Long

Dim oJSON As New jsonlib

Dim oCollection_sort As Collection
Set oCollection_sort = Me.ref_collection.transform_query_into_object(oQuery)

If oCollection_sort Is Nothing Then
    Debug.Print "WARNING cls_NoSQL_QueryResult::sort unable to get collection object from sort query"
    Exit Function
End If

Dim tmp_sort_filter As Collection
For Each tmp_sort_filter In oCollection_sort
    
    If tmp_sort_filter.count = 2 Then
    
        If VarType(tmp_sort_filter(1)) = vbString And VarType(tmp_sort_filter(2)) = vbDouble Then
            
            If tmp_sort_filter(2) = 1 Or tmp_sort_filter(2) = -1 Then
                
                's assure que champs dispo sinon supprime le document
                Dim tmp_vec_key() As Variant, tmp_vec_value() As Variant
                k = 0
                Dim tmp_oid As Variant
                Dim tmp_document As cls_NoSQL_Document
                For Each tmp_oid In Me.documents.keys
                    
                    Set tmp_document = Me.documents.Item(tmp_oid)
                    
                    If tmp_document.representation_dictionary.Exists(tmp_sort_filter(1)) = False Then
                        Me.documents.remove oJSON.toString(tmp_document.representation_dictionary.Item("_id"))
                    Else
                        ReDim Preserve tmp_vec_key(k)
                        tmp_vec_key(k) = oJSON.toString(tmp_document.representation_dictionary.Item("_id"))
                        
                        ReDim Preserve tmp_vec_value(k)
                        tmp_vec_value(k) = tmp_document.representation_dictionary.Item(tmp_sort_filter(1))
                        
                        k = k + 1
                    End If
                    
                Next
                
                
                Dim tmp_value As Variant
                Dim tmp_position As Long
                Dim tmp_key As Variant
                
                If tmp_sort_filter(2) = 1 Then
                    
                    For m = 0 To UBound(tmp_vec_value, 1)
                        
                        tmp_value = tmp_vec_value(m)
                        tmp_position = m
                        
                        For n = m + 1 To UBound(tmp_vec_value, 1)
                            
                            If tmp_vec_value(n) < tmp_value Then
                                tmp_value = tmp_vec_value(n)
                                tmp_position = n
                            End If
                            
                        Next n
                        
                        If tmp_position <> m Then
                            tmp_value = tmp_vec_value(m)
                            tmp_vec_value(m) = tmp_vec_value(tmp_position)
                            tmp_vec_value(tmp_position) = tmp_value
                            
                            tmp_key = tmp_vec_key(m)
                            tmp_vec_key(m) = tmp_vec_key(tmp_position)
                            tmp_vec_key(tmp_position) = tmp_key
                        End If
                        
                    Next m
                    
                ElseIf tmp_sort_filter(2) = -1 Then
                    
                    For m = 0 To UBound(tmp_vec_value, 1)
                        
                        tmp_value = tmp_vec_value(m)
                        tmp_position = m
                        
                        For n = m + 1 To UBound(tmp_vec_value, 1)
                            
                            If tmp_vec_value(n) > tmp_value Then
                                tmp_value = tmp_vec_value(n)
                                tmp_position = n
                            End If
                            
                        Next n
                        
                        If tmp_position <> m Then
                            tmp_value = tmp_vec_value(m)
                            tmp_vec_value(m) = tmp_vec_value(tmp_position)
                            tmp_vec_value(tmp_position) = tmp_value
                            
                            tmp_key = tmp_vec_key(m)
                            tmp_vec_key(m) = tmp_vec_key(tmp_position)
                            tmp_vec_key(tmp_position) = tmp_key
                        End If
                        
                    Next m
                    
                End If
                
                tmp_value = Empty
                
                Me.orders_keys = tmp_vec_key
                
            Else
                Debug.Print "WARNING cls_NoSQL_QueryResult::sort side must be 1 or -1"
                Exit Function
            End If
        
        Else
            
            Debug.Print "WARNING cls_NoSQL_QueryResult::sort query's datatype componenents incorrect, should be one string field and one number 1 / -1 for the side"
            Exit Function
        End If
    
    Else
        
        Debug.Print "WARNING cls_NoSQL_QueryResult::sort query's size incorrect, should be only 2 elements field and 1 (Ascending) or -1 (Descending)"
        Exit Function
    End If
    
    Exit For 'one level only pour l instant
Next

End Function


Public Function limit(qty As Long)



End Function


Public Function append_document(oDoc As cls_NoSQL_Document, Optional oid As Variant)

Dim oJSON As New jsonlib

Dim tmp_vec() As Variant
If Me.count = 0 Then
    ReDim Preserve tmp_vec(0)
Else
    tmp_vec = Me.orders_keys
    ReDim Preserve tmp_vec(Me.count)
End If

tmp_vec(UBound(tmp_vec, 1)) = oJSON.toString(oDoc.representation_dictionary.Item("_id"))

Me.orders_keys = tmp_vec

Me.documents.Add oJSON.toString(oDoc.representation_dictionary.Item("_id")), oDoc
Me.count = Me.count + 1

End Function


Private Sub Class_Initialize()

Set documents = New Scripting.Dictionary
count = 0

End Sub

